/**
 * クラウド同期API（合言葉方式）— 家計簿アプリの state を丸ごと1行として保存・取得する。
 * D1テーブル: budget_sync(passcode_hash text primary key, data text, updated_at text)
 *
 * GET  /api/budget?hash=<sha256hex>  -> { data, updated_at } / 404
 * POST /api/budget  { hash, data }   -> 保存（新規なら作成、既存なら上書き）
 */

const HASH_RE = /^[0-9a-f]{64}$/;

/** @param {{ request: Request, env: { DB: D1Database } }} ctx */
export async function onRequestGet({ request, env }) {
  const url = new URL(request.url);
  const hash = url.searchParams.get("hash") ?? "";
  if (!HASH_RE.test(hash)) {
    return new Response(JSON.stringify({ error: "invalid hash" }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }

  const row = await env.DB.prepare("select data, updated_at from budget_sync where passcode_hash = ?")
    .bind(hash)
    .first();

  if (!row) {
    return new Response(JSON.stringify({ error: "not found" }), {
      status: 404,
      headers: { "Content-Type": "application/json" },
    });
  }

  return new Response(
    JSON.stringify({ data: JSON.parse(/** @type {string} */ (row.data)), updated_at: row.updated_at }),
    { status: 200, headers: { "Content-Type": "application/json" } }
  );
}

/** @param {{ request: Request, env: { DB: D1Database } }} ctx */
export async function onRequestPost({ request, env }) {
  /** @type {{ hash?: unknown, data?: unknown }} */
  let body;
  try {
    body = await request.json();
  } catch {
    return new Response(JSON.stringify({ error: "invalid json" }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }

  const hash = typeof body.hash === "string" ? body.hash : "";
  if (!HASH_RE.test(hash)) {
    return new Response(JSON.stringify({ error: "invalid hash" }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }
  if (!body.data || typeof body.data !== "object") {
    return new Response(JSON.stringify({ error: "invalid data" }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }

  const dataText = JSON.stringify(body.data);
  const updatedAt = new Date().toISOString();

  await env.DB.prepare(
    "insert into budget_sync (passcode_hash, data, updated_at) values (?, ?, ?) " +
      "on conflict(passcode_hash) do update set data = excluded.data, updated_at = excluded.updated_at"
  )
    .bind(hash, dataText, updatedAt)
    .run();

  return new Response(JSON.stringify({ ok: true, updated_at: updatedAt }), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  });
}
