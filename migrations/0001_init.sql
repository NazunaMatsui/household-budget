-- 家計簿 クラウド同期用テーブル（合言葉のSHA-256ハッシュをキーに、アプリの state をJSON文字列で保持）
create table if not exists budget_sync (
  passcode_hash text primary key,
  data text not null,
  updated_at text not null
);
