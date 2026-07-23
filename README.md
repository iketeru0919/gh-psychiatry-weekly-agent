# 週刊 精神科・精神薬学・GH支援レポート生成MVP

PubMed E-utilities APIで精神科疾患・精神科薬学に関する新着論文を検索し、OpenAI APIで障がい者グループホーム（GH）の支援・リスク管理・職員研修に使いやすい論文をスコアリングして、日本語のMarkdownレポートを作成するPythonプロジェクトです。

## MVPでできること

- PubMed E-utilities APIから直近7日間の候補論文を最大100本取得します。
- 候補論文が少ない場合（既定値: 30本未満）は、直近14日間に自動拡大します。
- 取得項目:
  - PMID
  - タイトル
  - 著者
  - 雑誌名
  - 発行日
  - abstract
  - PubMed URL
- OpenAI APIで以下の観点を0〜5点で評価します。
  - GH現場との関連性
  - 精神科疾患理解への有用性
  - 精神科薬学・副作用理解への有用性
  - リスク管理への有用性
  - 職員研修への転用しやすさ
  - AM、管理者、サビ管が理解しやすいか
  - 研究の信頼性
- 最終的に10本程度を選び、日本語Markdownレポートを `reports/` に出力します。

## セットアップ

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

`.env` に以下を設定してください。

```env
OPENAI_API_KEY=sk-your-key
NCBI_EMAIL=you@example.com
```

NCBI_EMAILはNCBI E-utilities利用時の連絡先として使用します。実運用では必ず到達可能なメールアドレスを設定してください。

## 手動実行

```bash
PYTHONPATH=src python -m psychiatry_weekly_agent
```

特定日を検索終了日にする場合:

```bash
PYTHONPATH=src python -m psychiatry_weekly_agent --today 2026-05-06
```

実行後、`reports/weekly_report_YYYYMMDD_HHMMSS.md` が作成されます。

## レポート構成

出力されるMarkdownは以下の構成です。

1. 週刊 精神科・精神薬学・GH支援レポート
2. 対象期間
3. 今週の総括
4. 最重要3本
5. 疾患理解3本
6. 薬学・副作用2本
7. 支援・リスク管理・研修活用2本
8. 今週、職員へ共有するなら
9. 医学的判断を代替しない注意書き

各論文には、タイトル、分野、研究タイプ、重要度、要点、GH現場への示唆、AM・管理者・サビ管の確認ポイント、職員研修に使える一言、注意点、原文URLを含めます。

## GitHub Actionsでの週1回実行

`.github/workflows/weekly-report.yml` は、毎週月曜日 00:00 UTCにレポートを生成します。利用するにはリポジトリのSecretsに以下を登録してください。

- `OPENAI_API_KEY`
- `NCBI_EMAIL`

生成されたMarkdownはActions artifactとしてアップロードされます。

## 注意事項

このプロジェクトが生成するレポートは、PubMed掲載情報とabstractをもとにした教育・研修用の整理資料です。診断、治療、処方変更、服薬中止、緊急対応等の医学的判断を代替しません。利用者の症状悪化、副作用疑い、自傷他害リスク、急変がある場合は、主治医、薬剤師、訪問看護、救急、行政等の専門職・機関へ相談してください。

## おまけ: RASIEL 営業活動報告チャットの拠点別集計ツール

このリポジトリには、上記のPubMedレポート機能とは別に、グループホーム各拠点のLINE営業活動報告
チャットを拠点ごとに自動集計するツールも同梱しています。AI/外部APIを使わず、正規表現とラベル
辞書によるルールベース処理のみで動作します。

- `tools/rasiel_report_tool.html`: ブラウザで開いてログを貼り付けるだけで使える単一HTMLツール
- `src/rasiel_facility_reports/`: CSV集計・自動化用のPythonモジュール（`PYTHONPATH=src python -m rasiel_facility_reports --help`）

詳細は [`docs/rasiel_facility_report_proposal.md`](docs/rasiel_facility_report_proposal.md) を参照してください。
