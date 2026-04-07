# SV Portal 管理ガイド（Claude Code用）

このファイルを読んでいる Claude は SV Portal の保守・開発を担当しています。
**このファイルをまず全て読んでから作業を開始してください。**

---

## ポータル概要

美容サロンチェーンのSV（スーパーバイザー）が加盟店を管理するためのWebポータル。

- **フロントエンド**: `index.html`（GitHub Pages でホスティング）
- **バックエンド**: `gas_api.gs`（Google Apps Script Web App）
- **データストア**: Google スプレッドシート

ユーザーがポータルを開く → `index.html` が GAS API に GET リクエスト → GAS がスプレッドシートを読んで JSON を返す → ポータルに表示、という流れ。

---

## URL・ID 一覧

| 項目 | 値 |
|---|---|
| ポータルURL | `https://m-business-gif.github.io/sv-portal/` |
| GAS APIエンドポイント | `https://script.google.com/macros/s/AKfycbzWTeQ8iAgPJmEIlBnxQTb0GdNbeNgRq39nBa8vlcN4g_LADO6i1PAspsU2ocN_QITm/exec` |
| GASエディタ | `https://script.google.com/d/1FdMryXF5JaS4CkoKJSdzJvIOnplwLsVz-Agx3pXgd8CFADaFhtCo818H/edit` |
| スプレッドシートID | `1K-4ub8YvFh__JrseNKGiCkGigDYykraIwocOhLQLevY` |
| GitHubリポジトリ | `/Users/yamadamasaya/sv-portal`（ローカル） |

---

## ファイル構成と役割

```
sv-portal/
├── index.html       # ポータル全体（HTML+CSS+JavaScript が1ファイルに集約）
├── gas_api.gs       # GASバックエンド（スプレッドシート読み書き・API）
├── CLAUDE.md        # このファイル
└── appsscript.json  # GASプロジェクト設定（基本触らない）
```

### index.html の構成

`<script>` タグ内に全JavaScript。主な関数：

| 関数名 | 役割 |
|---|---|
| `loadData()` | GAS APIからデータ取得・SVC/TASK_CATSを設定シートで上書き |
| `render()` | 現在のタブに応じて表示を切り替える |
| `renderStores()` | 店舗一覧タブの描画 |
| `renderTasks()` | タスクボードタブの描画（カテゴリ別グループ） |
| `renderStaff()` | スタッフランクタブの描画 |
| `renderSum()` | サマリーカードの更新 |
| `openTaskModal(storeName, row)` | タスク追加・編集モーダルを開く（row=0で追加、row>0で編集） |
| `saveTaskModal()` | タスク保存（POST to GAS） |
| `changeStatus(row, newStatus, sel)` | タスクのステータスをその場で変更 |
| `autoFillSV()` | 店舗選択時にSVを自動入力 |

### gas_api.gs の構成

| 関数名 | 役割 |
|---|---|
| `doGet(e)` | GETリクエスト処理。stores/tasks/staffRanks/staffSales/configを返す |
| `doPost(e)` | POSTリクエスト処理。actionに応じて分岐 |
| `getStores()` | 加盟店目標・実数値・見込み数値シートから店舗データ集計 |
| `getTasks()` | タスクボードシートからタスク一覧取得 |
| `upsertTask(rowNum, taskData)` | タスク追加・更新（rowNum<2で新規追加） |
| `deleteTaskRow(rowNum)` | タスク削除 |
| `getConfig()` | 設定シートからSV一覧・タスクカテゴリ取得 |
| `setupConfig()` | 設定シートが空の場合にデフォルト値を書き込む |
| `setupTaskBoard()` | タスクボードシートの初期化 |
| `getStaffRanks()` | スタッフランクシート取得 |
| `getStaffSales()` | スタッフ売上シートから平均売上・ランク集計 |

---

## スプレッドシート構成

| シート名 | 役割 | 主な列 |
|---|---|---|
| `加盟店目標` | 店舗ごとの月次目標 | A:区分 B:SV C:店舗名 D:年月(YYYYMM) F:売上目標 G:新規 H:再来 |
| `実数値` | 店舗ごとの月次実績 | A:区分 B:? C:店舗名 D:年月 F:売上実績 G:新規 H:再来 |
| `見込み数値` | 月末着地見込み | C:店舗名 D:年月 J:総客数率 K:客単価 M:回数券 O:物販 |
| `タスクボード` | タスク管理 | A:店舗名 B:担当SV C:カテゴリ D:タスク名 E:ステータス F:優先度 G:メモ H:完了日 |
| `スタッフランク` | 手動設定のランク | A:店舗名 B:スタッフ名 C:ランク D:点数 |
| `売上明細（9~3月）` | 売上データ（自動集計元） | 多列（A:店舗 B:日付数値 F:区分 H:カテゴリ L:件数 M:金額 N:スタッフ名） |
| `【眉毛】加盟店管理集計` | **設定シート** | A:種別(SV/カテゴリ) B:値1(名前) C:値2(カラーコード) |

---

## よくある変更と手順

### SVを追加・変更する

**スプレッドシートだけで完結（コード不要）:**
1. `【眉毛】加盟店管理集計` シートを開く
2. `SV` 行を追加または編集する（例: `SV | 新SV名 | #カラーコード`）
3. ポータルをリロード → 反映完了

### タスクカテゴリを追加・変更する

**スプレッドシートだけで完結（コード不要）:**
1. `【眉毛】加盟店管理集計` シートを開く
2. `カテゴリ` 行を追加または編集する（例: `カテゴリ | 新カテゴリ名 |`）
3. ポータルをリロード → 反映完了

### 新しいタブ（画面）を追加する

1. `index.html` のナビゲーション部分に新しいタブボタンを追加
2. `render()` 関数に新しいタブの条件分岐を追加
3. 新しい `renderXxx()` 関数を実装
4. 必要ならGASに新しいデータ取得関数を追加し `doGet()` のレスポンスに含める

### 新しいデータ項目を表示する

1. GAS `getStores()` の `stores.push({...})` に新フィールドを追加
2. `index.html` の `renderStores()` でそのフィールドを参照して表示

### GASに新しいAPI操作を追加する

1. `gas_api.gs` の `doPost(e)` に新しい `if (data.action === "xxx")` ブランチを追加
2. 処理関数を実装
3. clasp push → deploy（下記参照）
4. `index.html` から `fetch(API_URL, {method:"POST", body: JSON.stringify({action:"xxx", ...})})` で呼び出す

### アラート基準を変更する

現在 `index.html` の `renderStores()` 内にハードコードされています：

```javascript
// 次回予約率アラート: 35%以下
// 回数券率アラート: 10%以下
// 施術単価アラート: 4800円未満
```

将来的には設定シートに移行予定。現時点ではコードを直接編集。

---

## デプロイ手順

### フロントエンド（index.html）の変更

```bash
cd /Users/yamadamasaya/sv-portal
git add index.html
git commit -m "変更内容の説明"
git push origin main
# → GitHub Pages が自動で反映（1〜2分）
```

### バックエンド（gas_api.gs）の変更

```bash
cd /Users/yamadamasaya/sv-portal
clasp push --force          # GASプロジェクトにファイルをアップロード
clasp deploy --description "説明"   # 新バージョンとしてデプロイ
```

**デプロイすると新しいURLが発行される。** その場合：
1. `index.html` の `API_URL` のデフォルト値を新URLに更新
2. このファイル（CLAUDE.md）の「GAS APIエンドポイント」を更新
3. git push

---

## データの流れ

```
ポータル（index.html）
    ↓ GET
GAS doGet()
    ↓ 読み取り
スプレッドシート（加盟店目標・実数値・見込み数値・タスクボード・設定）
    ↓ JSON返却
ポータル（stores, tasks, staffRanks, config を受け取り描画）

ポータル上でタスク操作
    ↓ POST {action: "upsertTask", ...}
GAS doPost()
    ↓ 書き込み
スプレッドシート（タスクボード）
```

---

## 重要なビジネスルール

- **フェリーチェ（`f`/`F` 系店舗）は加盟店分析から除外**して扱う
- **着地見込みは `見込み達成率` / `見込み売上` を使用**（`kgiForecast` は使わない）
- **売上データは内密扱い**。外部共有・不用意な出力はしない
- 分析アウトプットは「課題 → 店舗ごとのアクション → 全体でできるアクション」の3層構成

---

## 注意事項・地雷

- `index.html` は HTML/CSS/JS がすべて1ファイルに入っている。編集前に必ずRead toolで読む
- JavaScript変更後は必ず `node -e "const fs=require('fs');const code=fs.readFileSync('index.html','utf8');const m=code.match(/<script>([\s\S]*?)<\/script>/g);m.forEach((s,i)=>{try{new Function(s.replace(/<\/?script>/g,''));}catch(e){console.error('SYNTAX ERR',i,e.message);}});"` で構文チェック
- GAS再デプロイ後は `index.html` の `API_URL` も必ず更新すること（URLが変わるため）
- ローカルストレージに古いAPI_URLが残っていると新URLが反映されない（ユーザーに `localStorage.clear()` してもらう）
- `加盟店目標` シートの年月フィルタ: `dv >= 200001` の場合のみ当月YMと照合。`0` や空白は全月共通扱い

---

## API操作一覧（POST）

| action | 必須パラメータ | 説明 |
|---|---|---|
| `upsertTask` | `row`, `task` | タスク追加（row<2）または更新（row=行番号） |
| `deleteTask` | `row` | タスク削除 |
| `addStore` | `sv`, `name`, `ym`, `goals` | 新規加盟店追加 |
| `updateGoal` | `name`, `ym`, `goals` | 目標値更新 |
| `updateStaffRank` | `store`, `staff`, `rank`, `score` | スタッフランク手動設定 |
| `deleteStaff` | `store`, `staff` | スタッフ削除 |

---

## 現在の既知の課題・TODO

- アラート基準（次回予約率・回数券率・施術単価の閾値）が `index.html` にハードコード → 設定シートに移行予定
- タスク優先度（高/中/低）が `index.html` にハードコード → 設定シートに移行余地あり
