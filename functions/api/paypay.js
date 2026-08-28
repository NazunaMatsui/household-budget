/**
 * PayPayショートカット連携API — 既存のクラウド同期データ(budget_sync)に、支出エントリを1件追加する。
 * 家計簿アプリの「PayPay（食費）」カテゴリと同じ形（kind: expense, category固定）で追加される。
 *
 * POST /api/paypay  { hash, amount, note? }
 */

const HASH_RE = /^[0-9a-f]{64}$/;
const FOOD_CATEGORY = "PayPay（食費）";
const NOTE_MAX_LEN = 500;

function todayInTokyo() {
  // en-CA のロケール表記は YYYY-MM-DD になる
  return new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Tokyo" }).format(new Date());
}

/** @param {{ request: Request, env: { DB: D1Database } }} ctx */
export async function onRequestPost({ request, env }) {
  /** @type {{ hash?: unknown, amount?: unknown, note?: unknown }} */
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

  const amount = Math.floor(Number(body.amount));
  if (!Number.isFinite(amount) || amount <= 0) {
    return new Response(JSON.stringify({ error: "invalid amount" }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }

  const note = typeof body.note === "string" ? body.note.slice(0, NOTE_MAX_LEN) : "";

  const row = await env.DB.prepare("select data from budget_sync where passcode_hash = ?").bind(hash).first();
  if (!row) {
    return new Response(JSON.stringify({ error: "sync not set up for this passcode" }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }

  /** @type {{ entries?: unknown[], [k: string]: unknown }} */
  let data;
  try {
    data = JSON.parse(/** @type {string} */ (row.data));
  } catch {
    data = {};
  }
  const entries = Array.isArray(data.entries) ? data.entries : [];

  entries.push({
    id: crypto.randomUUID(),
    kind: "expense",
    date: todayInTokyo(),
    amount,
    category: FOOD_CATEGORY,
    note,
  });
  data.entries = entries;

  const updatedAt = new Date().toISOString();
  await env.DB.prepare("update budget_sync set data = ?, updated_at = ? where passcode_hash = ?")
    .bind(JSON.stringify(data), updatedAt, hash)
    .run();

  return new Response(JSON.stringify({ ok: true, amount, updated_at: updatedAt }), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  });
}
