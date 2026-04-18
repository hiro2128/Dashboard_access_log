# ALPFA Tableau Access Log Dashboard

Streamlit ダッシュボード。Tableau のアクセスログと社員情報を組み合わせて、利用状況を多角的に可視化します。

## 起動方法

```bash
streamlit run dashboard.py
```

## 入力ファイル

### データソースの優先順位

サイドバーのファイルアップローダーからファイルをアップロードすると、そのファイルが優先して使用されます。アップロードしない場合は、`dashboard.py` 冒頭の定数で指定したデフォルトパスのファイルが使用されます。

| 定数 | 説明 |
|------|------|
| `LOG_PATH` | アクセスログ Excel のデフォルトパス（Sheet1、2行目ヘッダー） |
| `USER_PATH` | ユーザー情報 Excel のデフォルトパス（"Personal User list" シート） |

現在のデフォルトパス:

```
LOG_PATH  = C:\Users\N206876\Documents\Tableau_aceess_log\CP-260413-1008.xlsx
USER_PATH = C:\Users\N206876\Documents\Tableau_aceess_log\【PROD】ALPFA Tableau users Information.xlsx
```

> **Note**: デフォルトパスのファイルが存在しない環境では、サイドバーから両ファイルをアップロードしてください。

### アクセスログの列構成（Sheet1、header=1）

| 列名 | 内容 |
|------|------|
| id | レコードID |
| project | 国 / プロジェクト |
| event_type | イベント種別 |
| timestamp | アクセス日時 |
| user | ユーザーID |
| workbook | ワークブック名 |
| view | ビュー（タブ）名 |

### ユーザー情報の列構成（Personal User list）

| 列名 | 内容 |
|------|------|
| First Name | 名 |
| Last Name | 姓 |
| User Name | ユーザーID（ログとの結合キー） |
| 所属 | 会社 / 所属組織 |
| Dept | 部署 |
| Status | アカウントステータス |

## 依存ライブラリ

```
streamlit
pandas
plotly
openpyxl   # pandas の Excel 読み込みに必要
```

## 画面構成

### サイドバー – ファイルアップロード

| 項目 | 説明 |
|------|------|
| Access Log Excel | アクセスログファイルのアップロード（任意）。未アップロード時はデフォルトパスを使用 |
| User Info Excel | ユーザー情報ファイルのアップロード（任意）。未アップロード時はデフォルトパスを使用 |

現在使用中のデータソース（アップロード済み or デフォルト）がサイドバーにバッジ表示されます。

---

### サイドバー – グローバルフィルター

全タブに共通して適用されるフィルター。

| フィルター | 対象列 |
|-----------|--------|
| Date Range | timestamp |
| Company | 所属（company） |
| Project Name | project |
| Workbook | workbook |
| View (Tab) | view |

フィルター後の件数と全体に対する割合をバッジで表示します。

---

### Tab 1 – Overview（📊）

| セクション | 内容 |
|-----------|------|
| KPI メトリクス | Total Access / Unique Users / Countries / Workbooks |
| Monthly Access Trend | 月次アクセス件数の棒グラフ |
| Access by Country | 上位 20 カ国の水平棒グラフ |
| Access by Workbook | 上位 15 ワークブックの水平棒グラフ |
| Weekday × Hour Heatmap | 曜日×時間帯のアクセス集中度（JST） |

---

### Tab 2 – Rankings（🏆）

| セクション | 内容 |
|-----------|------|
| Workbook Ranking | アクセス数順のワークブック一覧（順位・件数・構成比） |
| View (Tab) Ranking | アクセス数順のビュー一覧（順位・件数・構成比） |
| Country Ranking | 国別ランキング表 ＋ 上位 15 カ国のドーナツグラフ |

---

### Tab 3 – User Analysis（🔬）

| セクション | 内容 |
|-----------|------|
| Monthly Active Users | 月次ユニークユーザー数の折れ線グラフ |
| Heavy Users – Top 20 | アクセス数上位 20 名の表 ＋ 水平棒グラフ |
| User Deep Dive | ユーザー選択 → Dashboard Top 5 ＋ Project Top 5 のグラフ・表 |

---

### Tab 4 – Org Analysis（🏢）

| セクション | 内容 |
|-----------|------|
| Access by Company | 会社別ランキング表 ＋ 上位 12 社のドーナツグラフ |
| Access by Department | 上位 20 部署の水平棒グラフ |
| Company × Dashboard Usage Heatmap | 上位 15 社 × 上位 10 ダッシュボードの利用状況ヒートマップ |

## コード構成

```
dashboard.py
├── load_log()          # アクセスログ読み込み・派生列付与（Path または UploadedFile を受け取る）
├── load_users()        # ユーザー情報読み込み・整形（Path または UploadedFile を受け取る）
├── build_df_raw()      # ログとユーザー情報を結合して分析用 DataFrame を生成
├── make_ranking_df()   # ランキング表生成ヘルパー
├── add_rank_and_share()# 集計済み DataFrame に順位・構成比を付与するヘルパー
├── plot_hbar()         # 水平棒グラフ生成ヘルパー
├── plot_heatmap()      # ヒートマップ生成ヘルパー
├── サイドバー           # ファイルアップロード（任意）＋ グローバルフィルター
├── Tab 1 – Overview
├── Tab 2 – Rankings
├── Tab 3 – User Analysis
└── Tab 4 – Org Analysis
```
