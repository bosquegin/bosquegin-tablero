/**
 * bosquegin-actualizar — reemplazo de servidor_render.py (Flask/Render) por
 * un Cloudflare Worker sin servidor siempre-encendido ni cold-start.
 *
 * POST /actualizar  → valida cloud_token, chequea que no haya una corrida en
 *                      curso, dispara el workflow_dispatch de GitHub Actions.
 * GET  /estado       → estado de la corrida mas reciente del workflow
 *                      (para que el tablero haga polling en vez de mantener
 *                      una conexion abierta larga).
 * GET  /health        → chequeo simple.
 * POST /guardar_proyeccion_abastecimiento
 * POST /guardar_proyeccion_ingreso
 *                      → equivalentes a los endpoints homónimos de
 *                      servidor_bosquegin.py (local), pero para el sitio
 *                      publicado (sin backend propio): leen/escriben los
 *                      mismos archivos directo en GitHub via Contents API
 *                      y recalculan los campos derivados en JS (puerto de
 *                      _proy_recalcular_derivados de actualizar_bosquegin.py
 *                      — mismo cálculo, mismo resultado en los dos lugares).
 */

const ALLOWED_ORIGINS = new Set([
  "https://bosquegin.github.io",
  "https://bosquegin.com",
  "https://www.bosquegin.com",
]);

const AUTH_URL =
  "https://raw.githubusercontent.com/bosquegin/bosquegin-tablero/main/auth_static.js";

function corsHeaders(origin) {
  const allow = ALLOWED_ORIGINS.has(origin) ? origin : "";
  return {
    "Access-Control-Allow-Origin": allow,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Vary": "Origin",
  };
}

function json(body, status, origin) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders(origin) },
  });
}

async function fetchUsers() {
  const r = await fetch(AUTH_URL + "?t=" + Date.now(), { cf: { cacheTtl: 0 } });
  if (!r.ok) return [];
  const raw = await r.text();
  const m = raw.match(/window\.BG_AUTH\s*=\s*(\[[\s\S]*?\])\s*;/);
  if (!m) return [];
  try {
    return JSON.parse(m[1]);
  } catch (e) {
    return [];
  }
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let out = 0;
  for (let i = 0; i < a.length; i++) out |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return out === 0;
}

async function verifyToken(username, token) {
  if (!username || !token) return null;
  const users = await fetchUsers();
  for (const u of users) {
    if (u.username === username) {
      const stored = u.cloud_token || "";
      if (stored && timingSafeEqual(stored, token)) return u;
    }
  }
  return null;
}

function ghHeaders(env) {
  return {
    Authorization: `Bearer ${env.GH_TOKEN}`,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
    "User-Agent": "bosquegin-actualizar-worker",
  };
}

async function getRuns(env, params) {
  const url =
    `https://api.github.com/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}` +
    `/actions/workflows/${env.WORKFLOW_FILE}/runs?` +
    new URLSearchParams(params);
  const r = await fetch(url, { headers: ghHeaders(env) });
  if (!r.ok) throw new Error(`GitHub API runs ${r.status}`);
  return r.json();
}

async function dispatchWorkflow(env) {
  const url =
    `https://api.github.com/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}` +
    `/actions/workflows/${env.WORKFLOW_FILE}/dispatches`;
  const r = await fetch(url, {
    method: "POST",
    headers: { ...ghHeaders(env), "Content-Type": "application/json" },
    body: JSON.stringify({ ref: "main" }),
  });
  if (!r.ok) {
    const body = await r.text().catch(() => "");
    throw new Error(`GitHub API dispatch ${r.status}: ${body}`);
  }
}

// ═══════════════════════════════════════════════════════════════════════
//  GitHub Contents API — leer/escribir archivos individuales del repo
//  (base64 con soporte UTF-8, ya que atob/btoa solo manejan Latin1).
// ═══════════════════════════════════════════════════════════════════════

function b64EncodeUtf8(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

function b64DecodeUtf8(b64) {
  const bin = atob(b64.replace(/\n/g, ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

// Devuelve {content, sha} o null si el archivo no existe todavía.
async function ghGetFile(env, path) {
  const url =
    `https://api.github.com/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}` +
    `/contents/${path.split("/").map(encodeURIComponent).join("/")}`;
  const r = await fetch(url, { headers: ghHeaders(env) });
  if (r.status === 404) return null;
  if (!r.ok) throw new Error(`GitHub API get ${path} ${r.status}`);
  const data = await r.json();
  return { content: b64DecodeUtf8(data.content), sha: data.sha };
}

// sha=null crea el archivo si no existe; sha de un archivo existente lo
// actualiza (GitHub rechaza con 409 si el sha ya no es el vigente —
// alguien más lo cambió mientras tanto).
async function ghPutFile(env, path, content, sha, message) {
  const url =
    `https://api.github.com/repos/${env.GITHUB_OWNER}/${env.GITHUB_REPO}` +
    `/contents/${path.split("/").map(encodeURIComponent).join("/")}`;
  const body = { message, content: b64EncodeUtf8(content), branch: "main" };
  if (sha) body.sha = sha;
  const r = await fetch(url, {
    method: "PUT",
    headers: { ...ghHeaders(env), "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!r.ok) {
    const t = await r.text().catch(() => "");
    const err = new Error(`GitHub API put ${path} ${r.status}: ${t}`);
    err.status = r.status;
    throw err;
  }
  return r.json();
}

// ═══════════════════════════════════════════════════════════════════════
//  Proyección Producción — mismo cálculo que actualizar_bosquegin.py
//  (_proy_recalcular_derivados / aplicar_correcciones_abastecimiento /
//  aplicar_ingresos_abastecimiento), portado a JS para que editar desde el
//  sitio publicado dé exactamente el mismo resultado que desde el local.
// ═══════════════════════════════════════════════════════════════════════

const PROY_PREFIX =
  "window.BOSQUE_DATA=window.BOSQUE_DATA||{};window.BOSQUE_DATA.proyeccion=";
const PROY_CORR_PATH = "Data/Costos y PVP/Proyeccion_abastecimiento_correcciones.json";
const PROY_INGRESOS_PATH = "Data/Costos y PVP/Proyeccion_abastecimiento_ingresos.json";
const DATA_PROYECCION_PATH = "data_proyeccion.js";

function parseProyeccionJs(text) {
  const i = text.indexOf(PROY_PREFIX);
  if (i === -1) throw new Error("Formato de data_proyeccion.js inesperado");
  let body = text.slice(i + PROY_PREFIX.length).trim();
  if (body.endsWith(";")) body = body.slice(0, -1);
  return JSON.parse(body);
}

function serializeProyeccionJs(obj) {
  return PROY_PREFIX + JSON.stringify(obj) + ";";
}

async function readJsonFile(env, path) {
  const f = await ghGetFile(env, path);
  if (!f) return {};
  try {
    return JSON.parse(f.content);
  } catch (e) {
    return {};
  }
}

// Ver docstring de _proy_recalcular_derivados en actualizar_bosquegin.py
// para la explicación completa de cada regla — acá va el cálculo tal cual.
function proyRecalcularDerivados(p, mesesKeys) {
  let idxActual = -1;
  for (let i = 0; i < mesesKeys.length; i++) {
    const mm = p.mensual[mesesKeys[i]];
    if (mm && mm.proyeccion_mensual != null) {
      idxActual = i;
      break;
    }
  }
  const mesesRecalc = new Set(idxActual >= 0 ? mesesKeys.slice(idxActual) : mesesKeys);

  let saldoPrev = p.stock_actual;
  let totalObj = 0;
  let proyTotal = 0;
  for (const mk of mesesKeys) {
    const m = p.mensual[mk];
    if (!mesesRecalc.has(mk)) continue;
    const proy = m.ingreso ? 0 : (m.proyeccion_abastecimiento || 0);
    const vobj = m.venta_objetivo || 0;
    totalObj += vobj;
    proyTotal += proy;
    const salida = m.proyeccion_mensual != null ? m.proyeccion_mensual : vobj;
    saldoPrev = saldoPrev + proy - salida;
    m.saldo_stock = saldoPrev;
  }

  p.stock_total = p.stock_actual + proyTotal;
  p.total_objetivo_ventas = totalObj;
  p.meses_stock = totalObj > 0 ? p.stock_total / (totalObj / 3) : 0.0;
  const saldos = mesesKeys.map((mk) => p.mensual[mk].saldo_stock);
  p.comprar = Math.max(0, -Math.min(...saldos));
  p.alerta = p.comprar > 0 ? "COMPRAR" : "";
  p.cantidad_pallets = p.pallet ? Math.round((p.comprar / p.pallet) * 10) / 10 : 0.0;
}

// Formato de ambos archivos: {"<trimestre>": {"<cod>": {"<mesN>": valor}}}
function aplicarCorreccionesAbastecimiento(trimestres, correcciones) {
  for (const trimestre of Object.keys(correcciones || {})) {
    const T = trimestres[trimestre];
    if (!T) continue;
    const porCod = correcciones[trimestre];
    for (const p of T.productos || []) {
      const valores = porCod[String(p.cod)];
      if (!valores) continue;
      const mesesKeys = Object.keys(p.mensual || {});
      let cambio = false;
      for (const mk of Object.keys(valores)) {
        if (p.mensual[mk]) {
          p.mensual[mk].proyeccion_abastecimiento = Number(valores[mk]);
          cambio = true;
        }
      }
      if (cambio) proyRecalcularDerivados(p, mesesKeys);
    }
  }
}

function aplicarIngresosAbastecimiento(trimestres, ingresos) {
  for (const trimestre of Object.keys(ingresos || {})) {
    const T = trimestres[trimestre];
    if (!T) continue;
    const porCod = ingresos[trimestre];
    for (const p of T.productos || []) {
      const marcas = porCod[String(p.cod)];
      if (!marcas) continue;
      const mesesKeys = Object.keys(p.mensual || {});
      let cambio = false;
      for (const mk of Object.keys(marcas)) {
        if (p.mensual[mk]) {
          p.mensual[mk].ingreso = !!marcas[mk];
          cambio = true;
        }
      }
      if (cambio) proyRecalcularDerivados(p, mesesKeys);
    }
  }
}

// Aplica un edit (abastecimiento o ingreso) persistiéndolo primero en su
// archivo de correcciones/ingresos, y reaplica TODOS los guardados sobre
// data_proyeccion.js publicado (no solo el nuevo) — mismo patrón que
// servidor_bosquegin.py, para que un Actualizar posterior (que también
// reaplica desde estos mismos archivos) nunca pierda una edición manual.
// Reintenta una vez si el PUT de data_proyeccion.js choca por conflicto
// (409 — alguien más lo escribió mientras tanto).
async function guardarProyeccionEdit(env, { corrPath, trimestre, cod, mes, valor }) {
  let corrFile = await ghGetFile(env, corrPath);
  let datos = {};
  if (corrFile) {
    try {
      datos = JSON.parse(corrFile.content);
    } catch (e) {
      datos = {};
    }
  }
  datos[trimestre] = datos[trimestre] || {};
  datos[trimestre][cod] = datos[trimestre][cod] || {};
  datos[trimestre][cod][mes] = valor;
  await ghPutFile(
    env,
    corrPath,
    JSON.stringify(datos),
    corrFile ? corrFile.sha : null,
    `proyeccion: ${corrPath.includes("ingresos") ? "ingreso" : "abastecimiento"} ${cod} ${trimestre}/${mes}`
  );

  for (let intento = 0; intento < 2; intento++) {
    const proyFile = await ghGetFile(env, DATA_PROYECCION_PATH);
    if (!proyFile) throw new Error("No hay Proyección publicada todavía — correr Actualizar primero.");
    const cache = parseProyeccionJs(proyFile.content);

    const correcciones = await readJsonFile(env, PROY_CORR_PATH);
    const ingresos = await readJsonFile(env, PROY_INGRESOS_PATH);
    aplicarCorreccionesAbastecimiento(cache.trimestres || {}, correcciones);
    aplicarIngresosAbastecimiento(cache.trimestres || {}, ingresos);

    try {
      await ghPutFile(
        env,
        DATA_PROYECCION_PATH,
        serializeProyeccionJs(cache),
        proyFile.sha,
        `data: proyeccion abastecimiento ${cod} ${trimestre}/${mes}`
      );
      const T = (cache.trimestres || {})[trimestre];
      const producto = T ? (T.productos || []).find((p) => String(p.cod) === cod) : null;
      return producto;
    } catch (e) {
      if (e.status === 409 && intento === 0) continue; // reintentar una vez
      throw e;
    }
  }
}

async function handleGuardarAbastecimiento(request, env, origin) {
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return json({ ok: false, error: "JSON inválido" }, 400, origin);
  }
  const username = (body.username || "").trim();
  const token = (body.cloud_token || "").trim();
  const user = await verifyToken(username, token);
  if (!user) return json({ ok: false, error: "No autorizado" }, 401, origin);
  if (!["admin", "editor"].includes(user.role)) {
    return json({ ok: false, error: "Sin permiso" }, 403, origin);
  }

  const trimestre = String(body.trimestre || "");
  const cod = String(body.cod || "");
  const mes = String(body.mes || "");
  const valor = Number(body.valor);
  if (!trimestre || !cod || !mes || !isFinite(valor)) {
    return json({ ok: false, error: "Datos inválidos" }, 400, origin);
  }

  try {
    const producto = await guardarProyeccionEdit(env, {
      corrPath: PROY_CORR_PATH,
      trimestre,
      cod,
      mes,
      valor,
    });
    return json({ ok: true, producto }, 200, origin);
  } catch (e) {
    return json({ ok: false, error: String(e.message || e) }, 500, origin);
  }
}

async function handleGuardarIngreso(request, env, origin) {
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return json({ ok: false, error: "JSON inválido" }, 400, origin);
  }
  const username = (body.username || "").trim();
  const token = (body.cloud_token || "").trim();
  const user = await verifyToken(username, token);
  if (!user) return json({ ok: false, error: "No autorizado" }, 401, origin);
  if (!["admin", "editor"].includes(user.role)) {
    return json({ ok: false, error: "Sin permiso" }, 403, origin);
  }

  const trimestre = String(body.trimestre || "");
  const cod = String(body.cod || "");
  const mes = String(body.mes || "");
  const ingreso = !!body.ingreso;
  if (!trimestre || !cod || !mes) {
    return json({ ok: false, error: "Datos inválidos" }, 400, origin);
  }

  try {
    const producto = await guardarProyeccionEdit(env, {
      corrPath: PROY_INGRESOS_PATH,
      trimestre,
      cod,
      mes,
      valor: ingreso,
    });
    return json({ ok: true, producto }, 200, origin);
  } catch (e) {
    return json({ ok: false, error: String(e.message || e) }, 500, origin);
  }
}

async function handleActualizar(request, env, origin) {
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return json({ ok: false, error: "JSON inválido" }, 400, origin);
  }

  const username = (body.username || "").trim();
  const token = (body.cloud_token || "").trim();
  const user = await verifyToken(username, token);
  if (!user) return json({ ok: false, error: "No autorizado" }, 401, origin);
  if (!["admin", "editor"].includes(user.role)) {
    return json({ ok: false, error: "Sin permiso" }, 403, origin);
  }

  const data = await getRuns(env, { per_page: "5" });
  const runs = data.workflow_runs || [];
  const running = runs.some((r) => ["queued", "in_progress"].includes(r.status));
  if (running) {
    return json(
      { ok: false, error: "Ya hay una actualización en curso" },
      429,
      origin
    );
  }

  // Límite de 1 corrida por hora: cada Actualizar hace ~200 llamadas a la
  // API de Contabilium (stock por depósito + detalle de cada comprobante) y
  // dos corridas seguidas alcanzan para que Contabilium devuelva 429 y la
  // corrida quede a medias (ver incidente 2026-08-06).
  const COOLDOWN_MS = 60 * 60 * 1000;
  const lastRun = runs[0];
  if (lastRun) {
    const elapsedMs = Date.now() - new Date(lastRun.created_at).getTime();
    if (elapsedMs < COOLDOWN_MS) {
      const restanteMin = Math.ceil((COOLDOWN_MS - elapsedMs) / 60000);
      return json(
        {
          ok: false,
          error: `Actualizar tiene un límite de una vez por hora (para no saturar la API de Contabilium) — probá de nuevo en ${restanteMin} min.`,
        },
        429,
        origin
      );
    }
  }

  const dispatchedAt = Date.now();
  await dispatchWorkflow(env);

  return json({ ok: true, dispatchedAt }, 200, origin);
}

async function handleEstado(request, env, origin) {
  const url = new URL(request.url);
  const username = (url.searchParams.get("username") || "").trim();
  const token = (url.searchParams.get("cloud_token") || "").trim();
  const since = Number(url.searchParams.get("since") || "0");

  const user = await verifyToken(username, token);
  if (!user) return json({ ok: false, error: "No autorizado" }, 401, origin);

  const data = await getRuns(env, { per_page: "5" });
  const runs = data.workflow_runs || [];
  // La corrida que nos interesa: la más reciente creada después de "since"
  // (con un margen de 20s por si el reloj del cliente difiere), o si no hay
  // ninguna posterior, la más reciente en curso.
  let run =
    runs.find((r) => new Date(r.created_at).getTime() >= since - 20000) ||
    runs.find((r) => ["queued", "in_progress"].includes(r.status)) ||
    runs[0];

  if (!run) return json({ ok: true, status: "not_found" }, 200, origin);

  return json(
    {
      ok: true,
      status: run.status, // queued | in_progress | completed
      conclusion: run.conclusion, // success | failure | cancelled | null
      html_url: run.html_url,
      created_at: run.created_at,
    },
    200,
    origin
  );
}

export default {
  async fetch(request, env, ctx) {
    const origin = request.headers.get("Origin") || "";
    const url = new URL(request.url);

    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders(origin) });
    }

    try {
      if (url.pathname === "/health") {
        return json({ status: "ok", time: Date.now() }, 200, origin);
      }
      if (url.pathname === "/actualizar" && request.method === "POST") {
        return await handleActualizar(request, env, origin);
      }
      if (url.pathname === "/estado" && request.method === "GET") {
        return await handleEstado(request, env, origin);
      }
      if (url.pathname === "/guardar_proyeccion_abastecimiento" && request.method === "POST") {
        return await handleGuardarAbastecimiento(request, env, origin);
      }
      if (url.pathname === "/guardar_proyeccion_ingreso" && request.method === "POST") {
        return await handleGuardarIngreso(request, env, origin);
      }
      return json({ ok: false, error: "No encontrado" }, 404, origin);
    } catch (e) {
      return json({ ok: false, error: String(e.message || e) }, 500, origin);
    }
  },
};
