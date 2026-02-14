# PS1 SNOW Utilities

[日本語](#日本語) | [English](#english)

---

## 日本語

PS1 SNOW Utilities は、ServiceNow テーブルを CSV としてエクスポートするための PowerShell (WinForms) ユーティリティです。

### 基本的な使い方

1. `PS1SNOWUtilities.ps1` を実行します（PowerShell 5.1 / STA 推奨）。
2. **設定**タブで以下を入力します。
   - ServiceNow インスタンス名
   - 認証方式（ユーザID+パスワード または APIキー）
   - 必要に応じて UI 言語
3. **Export**タブで対象テーブルを選択します。
4. 必要に応じてフィルタ（全件 or `sys_updated_on` 期間指定）を設定します。
5. エクスポート先フォルダを指定して **実行** を押します。
6. ログを確認し、必要に応じて **フォルダを開く** で出力先を開きます。

### 免責事項

本ソフトウェアは ServiceNow 社とは無関係であり、ServiceNow 社による承認・保証・サポートを受けていません。

### ライセンス

本ソフトウェアは **MIT License** の下で提供されます。  
Copyright (c) ixam.net  
https://www.ixam.net

---

## English

PS1 SNOW Utilities is a PowerShell (WinForms) utility for exporting ServiceNow table data to CSV.

### Basic Usage

1. Run `PS1SNOWUtilities.ps1` (PowerShell 5.1 / STA recommended).
2. In the **Settings** tab, configure:
   - ServiceNow instance name
   - Authentication method (User ID + Password or API Key)
   - UI language if needed
3. In the **Export** tab, select the target table.
4. Optionally set filters (All records or `sys_updated_on` date range).
5. Choose an export directory and click **Execute**.
6. Check logs, and use **Open Folder** to view exported files.

### Disclaimer

This software is not affiliated with ServiceNow, and is not endorsed, supported, or warranted by ServiceNow.

### License

This software is licensed under the **MIT License**.  
Copyright (c) ixam.net  
https://www.ixam.net
