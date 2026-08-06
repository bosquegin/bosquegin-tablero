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
      return json({ ok: false, error: "No encontrado" }, 404, origin);
    } catch (e) {
      return json({ ok: false, error: String(e.message || e) }, 500, origin);
    }
  },
};
