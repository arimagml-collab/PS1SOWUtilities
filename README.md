# PS1 SNOW Utilities

[日本語](#日本語) | [English](#english)

---

## 日本語

PS1 SNOW Utilities は、ServiceNow テーブルを CSV / JSON / Excel (.xlsx) としてエクスポートするための PowerShell (WinForms) ユーティリティです。

### 基本的な使い方

1. `PS1SNOWUtilities.ps1` を実行します（PowerShell 5.1 / STA 推奨）。
2. **設定**タブで以下を入力します。
   - ServiceNow インスタンス名
   - 認証方式（ユーザID+パスワード または APIキー）
   - 必要に応じて UI 言語
3. **Export**タブで対象テーブルを選択します。
4. 必要に応じてフィルタ（全件 or `sys_updated_on` 期間指定）を設定します。
5. エクスポート先フォルダと出力形式（CSV / JSON / Excel）を指定して **実行** を押します。
6. ログを確認し、必要に応じて **フォルダを開く** で出力先を開きます。

### DataBase View Editor の制限

現行版の DataBase View Editor は、**ベーステーブル 1 つ** と **Where 句** の作成に対応しています。  
**JOIN / LEFT JOIN の定義は未対応**です。

### 免責事項

本ソフトウェアは ServiceNow 社とは無関係であり、ServiceNow 社による承認・保証・サポートを受けていません。

### ライセンス

本ソフトウェアは **MIT License** の下で提供されます。  
Copyright (c) ixam.net  
https://www.ixam.net

---

## English

PS1 SNOW Utilities is a PowerShell (WinForms) utility for exporting ServiceNow table data to CSV / JSON / Excel (.xlsx).

### Basic Usage

1. Run `PS1SNOWUtilities.ps1` (PowerShell 5.1 / STA recommended).
2. In the **Settings** tab, configure:
   - ServiceNow instance name
   - Authentication method (User ID + Password or API Key)
   - UI language if needed
3. In the **Export** tab, select the target table.
4. Optionally set filters (All records or `sys_updated_on` date range).
5. Choose an export directory and output format (CSV / JSON / Excel), then click **Execute**.
6. Check logs, and use **Open Folder** to view exported files.

### DataBase View Editor limitation

The current DataBase View Editor supports creating a view with **one base table** and a **where clause**.  
**JOIN / LEFT JOIN definitions are not supported** yet.

### Disclaimer

This software is not affiliated with ServiceNow, and is not endorsed, supported, or warranted by ServiceNow.

### License

This software is licensed under the **MIT License**.  
Copyright (c) ixam.net  
https://www.ixam.net
