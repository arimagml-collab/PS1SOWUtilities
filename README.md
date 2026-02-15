# PS1 SNOW Utilities

[日本語](#日本語) | [English](#english)

---

## 日本語

PS1 SNOW Utilities は、ServiceNow テーブルのデータ抽出（Export）と Database View の作成（Database View Editor）を行える PowerShell (WinForms) ユーティリティです。

### タブ別の活用シーン

- **Export**
  - データを CSV / JSON / Excel に出力し、各部署で自由に集計・加工・連携したいときに有効です。
  - 例：運用部門が Excel で一次分析し、別チームが JSON を使って別システム連携する並行利用。
- **Database View Editor**
  - ServiceNow 標準 UI では操作しづらい Database View 作成を、GUI で手早く組み立てたいときに有効です。
  - テーブル/カラム候補を見ながら、ベーステーブルと JOIN を設計できます。
- **設定**
  - インスタンス名・認証方式・言語などを保存し、繰り返し作業の入力ミスやセットアップ時間を減らしたいときに有効です。

### 前提条件

- Windows + PowerShell 5.1（WinForms 利用のため）
- ServiceNow インスタンスにアクセスできるネットワーク
- 対象テーブル参照権限（Export）および Database View 作成に必要な権限（View Editor）

### 基本的な使い方

1. `PS1SNOWUtilities.ps1` を実行します（PowerShell 5.1 / STA 推奨）。
2. **設定**タブで以下を入力します。
   - ServiceNow インスタンス名
   - 認証方式（ユーザID+パスワード または APIキー）
   - 必要に応じて UI 言語
3. 入力内容は `settings.json` に自動保存されます。

#### Export の手順

1. **Export**タブで対象テーブルを選択（または手動入力）します。
2. 必要に応じてフィルタ（全件 or `sys_updated_on` 期間指定）を設定します。
3. エクスポート先フォルダと出力形式（CSV / JSON / Excel）を指定して **実行** を押します。
4. ログを確認し、必要に応じて **フォルダを開く** で出力先を開きます。

#### Database View Editor の手順

1. **Database View Editor** タブで View 内部名と View ラベルを入力します。
2. ベーステーブルを選択し、必要に応じてベース Prefix を設定します。
3. **JOIN追加** で JOIN テーブル・左右カラム・Variable Prefix・LEFT JOIN 条件を設定します。
4. **カラム再取得** でカラム候補を再読み込みします（現状は候補がそのまま表示カラムとして扱われます）。
5. **View作成** を実行し、完了ログとリンク（作成済み View 一覧 / View 定義）を確認します。

### 補足（権限・制約）

- テーブル一覧は `sys_db_object` から取得するため、ACL により一覧取得できない場合があります（その場合は手動入力で対応）。
- 環境によっては Where 句または JOIN 定義の自動保存に制約があり、View 本体作成後に ServiceNow 側で手動補完が必要な場合があります。

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
また、作成者自身も本ソフトウェアの利用により生じたいかなる損害についても責任を負いません。ご利用にあたっては、必ず利用者自身の責任で十分にテストと確認を行ったうえでご活用ください。

### ライセンス

本ソフトウェアは **MIT License** の下で提供されます。  
Copyright (c) ixam.net  
https://www.ixam.net

---

## English

PS1 SNOW Utilities is a PowerShell (WinForms) utility for exporting ServiceNow table data and creating Database Views with a guided GUI.

### Useful situations by tab

- **Export**
  - Best when you want to distribute data as CSV / JSON / Excel so each department can process it in its own workflow.
  - Example: the operations team analyzes in Excel while another team consumes JSON for system integration.
- **Database View Editor**
  - Best when ServiceNow's native UI feels cumbersome for building Database Views.
  - You can design base tables and joins while checking table/column candidates.
- **Settings**
  - Best when you want to persist instance/auth/language preferences and reduce repeated setup time and input mistakes.

### Prerequisites

- Windows + PowerShell 5.1 (WinForms-based UI)
- Network access to your ServiceNow instance
- Appropriate permissions for table reads (Export) and Database View creation (View Editor)

### Basic Usage

1. Run `PS1SNOWUtilities.ps1` (PowerShell 5.1 / STA recommended).
2. In the **Settings** tab, configure:
   - ServiceNow instance name
   - Authentication method (User ID + Password or API Key)
   - UI language if needed
3. Inputs are auto-saved to `settings.json`.

#### Export workflow

1. In the **Export** tab, select the target table (or type it manually).
2. Optionally set filters (All records or `sys_updated_on` date range).
3. Choose an export directory and output format (CSV / JSON / Excel), then click **Execute**.
4. Check logs and use **Open Folder** to view exported files.

#### Database View Editor workflow

1. In the **Database View Editor** tab, enter the View name and label.
2. Select a base table, and set the base prefix if required.
3. Use **Add Join** to define join table, left/right columns, variable prefix, and LEFT JOIN options.
4. Click **Reload Columns** to refresh column candidates (currently, the loaded candidates are treated as visible columns as-is).
5. Click **Create View**, then review completion logs and links (created View list / View definition record).

### Notes (permissions and limitations)

- The table list is retrieved from `sys_db_object`; if blocked by ACL, enter table names manually.
- Depending on your instance, automatic persistence of where clause or join definitions may be limited. In that case, complete them manually in ServiceNow after the View itself is created.

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
The author also accepts no liability for any damages arising from the use of this software. You are responsible for thoroughly testing and verifying it before use.

### License

This software is licensed under the **MIT License**.  
Copyright (c) ixam.net  
https://www.ixam.net

### Images
<img width="1106" height="713" alt="snow_util_01" src="https://github.com/user-attachments/assets/1eea1cf8-c8b2-4a61-a71d-387daa5a8513" />
<img width="1106" height="713" alt="snow_util_02" src="https://github.com/user-attachments/assets/8b73fb3e-fede-45a3-96fa-4bdee30567fc" />
<img width="1106" height="713" alt="snow_util_03" src="https://github.com/user-attachments/assets/242a2530-b023-437f-8866-95f226f42d52" />

