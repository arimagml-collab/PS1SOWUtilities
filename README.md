# PS1 SNOW Utilities

[日本語](#日本語) | [English](#english)

---

## 日本語

PS1 SNOW Utilities は、ServiceNow テーブルのデータ抽出（Export）と DataBase View の作成（DataBase View Editor）を行える PowerShell (WinForms) ユーティリティです。

### タブ別の活用シーン

- **Export**
  - データを CSV / JSON / Excel に出力し、各部署で自由に集計・加工・連携したいときに有効です。
  - 例：運用部門が Excel で一次分析、別チームが JSON を使って別システム連携するといった並行利用。
- **DataBase View Editor**
  - ServiceNow 標準 UI では操作しづらい Database View 作成を、GUI で手早く組み立てたいときに有効です。
  - Admin 権限がなくテーブル/カラムの内部名を把握しづらい場合でも、候補を見ながらベーステーブル・JOIN・表示カラムを作成できます。
- **設定**
  - インスタンス名・認証方式・言語などを保存し、繰り返し作業の入力ミスやセットアップ時間を減らしたいときに有効です。

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

### 認証情報の保存方式（パスワード / APIキー）

- `settings.json` に保存される `passwordEnc` / `apiKeyEnc` は、Windows の **DPAPI (CurrentUser)** で暗号化されています。
- そのため、通常は **同じ Windows ユーザー + 同じ PC** でのみ復号でき、別PCへ `settings.json` をコピーしても読み取りできません。
- 復号キーをレジストリへ別保存する実装は採用していません（レジストリ依存なし）。
- より厳格にしたい場合は、次の運用を推奨します。
  - APIキーは短寿命トークン化・定期ローテーションする
  - 端末移行時は `settings.json` の秘密情報を引き継がず再入力する
  - 企業環境では Windows Credential Manager / SecretManagement 連携を検討する

### 免責事項

本ソフトウェアは ServiceNow 社とは無関係であり、ServiceNow 社による承認・保証・サポートを受けていません。

### ライセンス

本ソフトウェアは **MIT License** の下で提供されます。  
Copyright (c) ixam.net  
https://www.ixam.net

---

## English

PS1 SNOW Utilities is a PowerShell (WinForms) utility for both exporting ServiceNow table data and creating DataBase Views more comfortably.

### Useful situations by tab

- **Export**
  - Best when you want to distribute data as CSV / JSON / Excel so each department can process it in its own workflow.
  - Example: operations team analyzes in Excel while another team consumes JSON for system integration.
- **DataBase View Editor**
  - Best when ServiceNow's native UI feels cumbersome for building Database Views.
  - Especially helpful without admin privileges, where internal table/column names are hard to find; you can still build base tables, joins, and visible columns efficiently.
- **Settings**
  - Best when you want to persist instance/auth/language preferences and reduce repeated setup time and input mistakes.

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

### Credential storage model (Password / API Key)

- `passwordEnc` and `apiKeyEnc` in `settings.json` are encrypted with Windows **DPAPI (CurrentUser)**.
- In normal use, secrets can be decrypted only by the **same Windows user on the same machine**. Copying `settings.json` to another PC should not make secrets readable.
- This project does not rely on a separate registry-stored decryption key.
- For stricter operations, consider:
  - Short-lived API tokens with regular rotation
  - Re-entering secrets after device migration instead of carrying encrypted blobs
  - Enterprise-backed secret stores (Windows Credential Manager / SecretManagement)

### Disclaimer

This software is not affiliated with ServiceNow, and is not endorsed, supported, or warranted by ServiceNow.

### License

This software is licensed under the **MIT License**.  
Copyright (c) ixam.net  
https://www.ixam.net
