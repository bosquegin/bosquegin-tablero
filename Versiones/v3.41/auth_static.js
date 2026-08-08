// Deprecado (v3.41) — el login y el admin de usuarios ahora se validan
// server-side contra Cloudflare Workers KV, no contra este archivo.
// Antes acá vivían password_hash/salt/cloud_token de cada usuario, en un
// archivo del repo PÚBLICO -- ver v3.41 en CHANGELOG.md. Se deja este
// stub (en vez de borrar el archivo) para que el historial de commits de
// versiones viejas no encuentre un archivo faltante al compararlas.
window.BG_AUTH = [];
