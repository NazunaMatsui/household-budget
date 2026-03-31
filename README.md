# 家計簿（PWA）

## `http://localhost` で開く

### 方法 A: Python 3（Node 不要）

ターミナルでこのフォルダへ移動してから:

```bash
python3 serve-local.py
```

ブラウザで **http://localhost:5173** を開きます。

### 方法 B: npm（`serve`）

```bash
npm install
npm start
```

同じく **http://localhost:5173** です。

---

PWA（ホーム画面への追加など）は **`https` または `http://localhost`** で開いたときに利用できます。`file://` では制限があります。

停止するときはターミナルで `Ctrl+C` です。
