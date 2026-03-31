# SV Portal 管理ガイド（Claude Code用）

このファイルを読んでいる Claude は、SV Portal の管理操作を担当しています。
以下の手順に従って操作してください。

---

## GAS API エンドポイント

```
https://script.google.com/macros/s/AKfycbyuKP5MEYwmvgsV2kaUUPpsJCk_biBO0qLGYGn2GTRubRQEE-HVXvjIL3rfuyHU1Lzp/exec
```

このURLへのPOSTで全操作が完結します。スプレッドシートへの直接アクセスは不要です。

---

## スプレッドシート（参照用）

- **GASバックエンド用**: `1K-4ub8YvFh__JrseNKGiCkGigDYykraIwocOhLQLevY`
  - シート: `加盟店目標` / `実数値` / `見込み数値` / `タスクボード` / `スタッフランク`

---

## 操作一覧

### 1. 現在のデータ確認（GET）

```bash
curl "https://script.google.com/macros/s/AKfycbyuKP5MEYwmvgsV2kaUUPpsJCk_biBO0qLGYGn2GTRubRQEE-HVXvjIL3rfuyHU1Lzp/exec" -L
```

`stores` / `tasks` / `staffRanks` が返ります。

---

### 2. 店舗追加（addStore）

新しい加盟店を追加する場合:

```bash
curl -X POST "https://script.google.com/macros/s/AKfycbyuKP5MEYwmvgsV2kaUUPpsJCk_biBO0qLGYGn2GTRubRQEE-HVXvjIL3rfuyHU1Lzp/exec" \
  -H "Content-Type: application/json" \
  -d '{
    "action": "addStore",
    "sv": "山田",
    "name": "S新宿2号店",
    "ym": "202604",
    "goals": {
      "sales": 1500000,
      "newGuest": 80,
      "repeat": 120,
      "total": 200,
      "unitPrice": 7500,
      "royalty": 120000,
      "wholesale": 50000,
      "svSales": 170000
    }
  }'
```

**SVの選択肢**: `山田` / `髙橋` / `向井` / `子龍` / `宮脇`

**goalsのフィールド**:
| キー | 意味 |
|---|---|
| `sales` | 店舗売上目標 |
| `newGuest` | 新規客数目標 |
| `repeat` | 再来客数目標 |
| `total` | 総客数目標 |
| `unitPrice` | 客単価目標 |
| `ticket` | 回数券売上目標 |
| `royalty` | ロイヤリティ売上目標 |
| `wholesale` | 卸売上目標 |
| `svSales` | SV売上目標 |

---

### 3. 目標値の更新（updateGoal）

既存店舗の月次目標を変更する場合（行が無ければ自動追加）:

```bash
curl -X POST "https://script.google.com/macros/s/AKfycbyuKP5MEYwmvgsV2kaUUPpsJCk_biBO0qLGYGn2GTRubRQEE-HVXvjIL3rfuyHU1Lzp/exec" \
  -H "Content-Type: application/json" \
  -d '{
    "action": "updateGoal",
    "name": "S池袋",
    "ym": "202605",
    "goals": {
      "sales": 1600000,
      "newGuest": 85
    }
  }'
```

変更したいフィールドだけ `goals` に含めればOK（他フィールドは上書きされません）。

---

### 4. タスクの追加・更新（upsertTask）

```bash
curl -X POST "..." \
  -H "Content-Type: application/json" \
  -d '{
    "action": "upsertTask",
    "row": -1,
    "task": {
      "store": "S池袋",
      "sv": "山田",
      "category": "集客",
      "name": "HPBクーポン更新",
      "status": "未着手",
      "priority": "高",
      "memo": "5月末期限"
    }
  }'
```

`row: -1` で新規追加。既存行を更新する場合は行番号を指定。

---

### 5. タスク削除（deleteTask）

```bash
curl -X POST "..." \
  -H "Content-Type: application/json" \
  -d '{ "action": "deleteTask", "row": 5 }'
```

---

## 年月（YM）フォーマット

`YYYYMM` 形式。例: `202604` = 2026年4月

---

## 注意事項

- フェリーチェ（`f`/`F`系店舗）は加盟店分析から除外して扱う
- 売上データは内密扱い。不用意な共有・出力はしない
- GAS APIはDeployされたウェブアプリのURL。スクリプトを変更した場合は**再デプロイが必要**

## GAS 更新・再デプロイ手順（clasp）

```bash
cd /Users/yamadamasaya/sv-portal
clasp push                           # gas_api.gs をGASプロジェクトに反映
clasp deploy --description "説明"    # 新バージョンとしてデプロイ → 新URLが発行される
```

デプロイ後、URLが変わった場合は `index.html` の `API_URL` と `CLAUDE.md` のURLを更新してgit push。

**GASエディタURL（権限承認・デバッグ用）:**
```
https://script.google.com/d/1FdMryXF5JaS4CkoKJSdzJvIOnplwLsVz-Agx3pXgd8CFADaFhtCo818H/edit
```

---

## ポータルURL

```
https://m-business-gif.github.io/sv-portal/
```

---

## リポジトリ構成

| ファイル | 役割 |
|---|---|
| `index.html` | ポータルのフロントエンド（GitHub Pages） |
| `gas_api.gs` | GASバックエンド（Google Apps Script） |
| `CLAUDE.md` | このファイル：Claude Code向け操作ガイド |
