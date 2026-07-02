# 請求提供差異チェック

## 目的

請求ExcelとDropbox内の提供実績Excelを読み取り、以下を自動チェックします。

- DBシートの請求人数・提供人数
- 請求側の生活援助日中系回数
- 提供側の延べ日数
- 利用者別の差異
- 差異理由候補、根拠、DB Q列への記載案

## 通常の使い方

`run_check_double_click.bat` をダブルクリックしてください。

起動後、画面で対象年・対象月を入力します。

```text
Target year [2026]
Target month 1-12 [5]
```

Enterだけ押すと、`config_auto.json` の既定値を使います。

既定では `config_auto.json` の設定を使い、以下の出力ファイルを作成します。

```text
billing_provider_check_YYYY_MM.xlsx
```

例：

```text
billing_provider_check_2026_05.xlsx
```

請求Excelをこのフォルダに入れておくと、起動時に候補として表示されます。
候補がない場合は、`config_auto.json` の `billing_workbook` を使います。

## 年月を指定して実行する場合

PowerShellで以下のように実行します。

```powershell
powershell -ExecutionPolicy Bypass -File ".\run_check.ps1" -Year 2026 -Month 5
```

6月を確認する場合は以下です。

```powershell
powershell -ExecutionPolicy Bypass -File ".\run_check.ps1" -Year 2026 -Month 6
```

## 設定ファイル

通常運用では `config_auto.json` を編集します。

主な項目は以下です。

```json
{
  "target_year": 2026,
  "target_month": 5,
  "billing_workbook": "C:/.../【5月実績】ラシエル稼働_202626.xlsx",
  "provider_root": "C:/Users/k-takayama/Bihonest Corp. Dropbox/.../実績記録表　保存先",
  "facilities": [
    "南中丸"
  ]
}
```

`provider_root` には、施設名別フォルダが格納されているDropbox上の親フォルダを指定します。

## Dropbox自動探索

プログラムは以下の順で提供実績Excelを探します。

```text
provider_root
  ↓
施設フォルダ、例 RASIEL南中丸
  ↓
請求用
  ↓
2026年
  ↓
2026年5月提供分
  ↓
*.xlsm / *.xlsx
```

南中丸のように `1F` と `2F` がある場合は、自動で合算します。

入居系チェックでは、ファイル名に `短期` を含むExcelは自動除外します。

## 出力Excelで見るシート

まず以下を見てください。

- `サマリー`
- `差異一覧`
- `読取ファイル一覧`

`読取ファイル一覧` では、DropboxからどのExcelを読んだか確認できます。

## 判定

- `OK`: 主要な人数・延べ日数が一致
- `注意`: 人数またはDB T列だけ乖離しているが、請求・提供の延べ日数は一致（DB行未検出の場合も注意）
- `要確認`: 利用者別または延べ日数に差異あり
- `重大`: 提供あり請求なし、または請求あり提供なしの可能性あり

## 廃止ファイル

`build_report.mjs` は旧環境専用の別実装（`@oai/artifact-tool` 依存で通常環境では動作しない）のため廃止しました。
レポート生成は `build_report.py` に一本化しています。
ローカルのフォルダに `build_report.mjs` が残っている場合は削除してください。
