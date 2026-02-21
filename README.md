# PS1 SNOW Utilities

[日本語](#日本語) | [English](#english)

---

## 日本語

PS1 SNOW Utilities は、ServiceNow テーブルのデータ抽出（Export）・添付ファイル回収（Attachment Harvester）・Database View の作成（Database View Editor）・レコード全削除（Truncate）を行える PowerShell (WinForms) ユーティリティです。

### タブ別の活用シーン

- **Export**
  - データを CSV / JSON / Excel に出力し、各部署で自由に集計・加工・連携したいときに有効です。
  - 例：運用部門が Excel で一次分析し、別チームが JSON を使って別システム連携する並行利用。
- **Attachment Harvester**
  - 指定期間内に更新されたレコードに紐づく添付ファイルを一括取得したいときに有効です。
  - ファイル名は「テーブル名_レコードキー(number/short_description/sys_id)_元ファイル名」形式で保存され、重複時は連番で衝突回避します。
- **Database View Editor**
  - ServiceNow 標準 UI では操作しづらい Database View 作成を、GUI で手早く組み立てたいときに有効です。
  - テーブル/カラム候補を見ながら、ベーステーブルと JOIN を設計できます。
- **Truncate（全削除）**
  - 開発環境で数万オーダーの大量データインポートテストを繰り返し、テーブル管理画面からのレコード削除では大変な場合に有効です。
  - **本番環境での使用は非推奨**です。
- **設定**
  - インスタンス名・認証方式・言語などを保存し、繰り返し作業の入力ミスやセットアップ時間を減らしたいときに有効です。

### 前提条件

- Windows + PowerShell 5.1（WinForms 利用のため）
- ServiceNow インスタンスにアクセスできるネットワーク
- 対象テーブル参照権限（Export / Attachment Harvester）および Database View 作成に必要な権限（View Editor）
- 添付ファイル取得のため `sys_attachment` / 添付バイナリ API にアクセスできる権限（Attachment Harvester）

### 基本的な使い方

1. `PS1SNOWUtilities.ps1` を実行します（PowerShell 5.1 / STA 推奨）。
2. **設定**タブで以下を入力します。
   - ServiceNow インスタンス名
   - 認証方式（ユーザID+パスワード または APIキー）
   - 必要に応じて UI 言語
3. 入力内容はアプリ初回実行後に生成される `settings.json` に自動保存されます（リポジトリには含めていません）。

##### 独自ドメイン運用時の設定（`instance-name.service-now.com` 以外）

`settings.json` に `instanceDomain` を追加すると、API 接続先 URL を明示指定できます。

```json
{
  "instanceName": "dev12345",
  "instanceDomain": "example.com"
}
```

または、`instanceName` を空欄にして `instanceDomain` に `https://` から始まるフル URL を設定することもできます。

```json
{
  "instanceName": "",
  "instanceDomain": "https://example.com"
}
```

- `instanceDomain` を設定した場合はそちらが優先されます。
- `instanceDomain` には `example.com` または `https://example.com` のどちらでも指定できます（`https://` なしで記載した場合は自動補完）。
- `instanceDomain` が未設定または空の場合は、従来どおり `instanceName` から `https://<instanceName>.service-now.com` を組み立てます。

#### Export の手順

1. **Export**タブで対象テーブルを選択（または手動入力）します。
2. 必要に応じてフィルタ（全件 or `sys_updated_on` 期間指定）を設定します。
3. エクスポート先フォルダと出力形式（CSV / JSON / Excel）を指定して **実行** を押します。
4. ログを確認し、必要に応じて **フォルダを開く** で出力先を開きます。

##### CSV分割エクスポートの使い方

1. 出力形式で **CSV** を選択します。
2. **CSV分割エクスポート** を有効にし、1ファイルあたりの分割件数（行数）を指定します。
3. 実行すると、連番付きの複数CSVファイルとして順次出力されます。
4. ログで各ファイルの出力状況を確認し、必要に応じて後続処理で結合・集計します。

> 💡 使用シチュエーション：巨大テーブルを1ファイルで出力すると、ネットワークや処理時間の都合で途中で切れてしまう可能性がある場合に、分割して全件を安全に出力したいときに有効です。

#### Attachment Harvester の手順

1. **Attachment Harvester** タブで対象テーブルを選択（または手動入力）します。
2. 判定対象の日付項目（例：`sys_updated_on`）と期間（開始・終了日時）を指定します。
3. ダウンロード先フォルダを指定し、必要に応じて「テーブルごとにサブフォルダ作成」を有効化します。
4. 実行すると、期間条件に一致したレコードに紐づく添付ファイルを取得し、重複内容はハッシュ比較でスキップします。
5. ログで保存件数/スキップ件数/失敗件数を確認します。

> 💡 使用シチュエーション：障害調査や監査対応で、特定期間に更新されたチケットの証跡ファイルをまとめて回収したいときに有効です。

#### Database View Editor の手順

1. **Database View Editor** タブで View 内部名と View ラベルを入力します。
2. ベーステーブルを選択し、必要に応じてベース Prefix を設定します。
3. **JOIN追加** で JOIN テーブル・左右カラム・Variable Prefix・LEFT JOIN 条件を設定します。
4. **カラム再取得** でカラム候補を再読み込みします（現状は候補がそのまま表示カラムとして扱われます）。
5. **View作成** を実行し、完了ログとリンク（作成済み View 一覧 / View 定義）を確認します。

#### Truncate（全削除）の手順

1. **Truncate（全削除）**タブで削除対象テーブルを選択（または手動入力）します。
2. Truncate許可インスタンス（読取専用）を確認します。編集は UI では行わず、`settings.json` の `truncateAllowedInstances` を直接編集してください（ワイルドカード指定・カンマ区切り複数指定可、既定値: `*dev*,*stg*`）。
3. 最大再試行回数（1～999）を設定し、**全件削除実行** を押します。
4. 表示される確認コード入力ダイアログで4文字コードを入力し、実行確認ダイアログで承認します。
5. 進捗バーとログを確認し、必要に応じて再試行ログ（最大再試行回数まで）を確認します。

> 💡 使用シチュエーション：開発環境で数万オーダーの大量データインポートテストを繰り返すために削除したいが、テーブル管理画面から実施するレコード削除では大変な時に使用します。

> ⚠️ 本機能は本番環境での利用を推奨しません。開発環境での大量データインポート試験など、限定的な用途でのみ利用してください。

#### 社内配布向けに特定機能を除外する方法

- 危険性のある機能（例：Truncate）を含めずに配布したい場合は、`modules/Features` 配下の該当機能ファイル（例：`TruncateFeature.psm1`）を配布対象から除外してください。
- 除外した機能はアプリ起動時に読み込まれないため、対応タブ/操作は UI に表示されません。
- これにより、同一コードベースでも配布用途に応じて機能を絞った構成にできます。

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

PS1 SNOW Utilities is a PowerShell (WinForms) utility for exporting ServiceNow table data, harvesting attachments, creating Database Views, and truncating table records with a guided GUI.

### Useful situations by tab

- **Export**
  - Best when you want to distribute data as CSV / JSON / Excel so each department can process it in its own workflow.
  - Example: the operations team analyzes in Excel while another team consumes JSON for system integration.
- **Attachment Harvester**
  - Useful when you need to bulk-download attachments linked to records updated within a specific time window.
  - Files are saved as `table_recordKey(number/short_description/sys_id)_originalFileName`, and duplicate names are safely suffixed.
- **Database View Editor**
  - Best when ServiceNow's native UI feels cumbersome for building Database Views.
  - You can design base tables and joins while checking table/column candidates.
- **Truncate (Delete all)**
  - Useful when you repeatedly run large-volume import tests (tens of thousands of records) in development, and deleting records from the table management screen is too cumbersome.
  - **Not recommended for production use**.
- **Settings**
  - Best when you want to persist instance/auth/language preferences and reduce repeated setup time and input mistakes.

### Prerequisites

- Windows + PowerShell 5.1 (WinForms-based UI)
- Network access to your ServiceNow instance
- Appropriate permissions for table reads (Export / Attachment Harvester) and Database View creation (View Editor)
- Access to `sys_attachment` and attachment binary APIs for downloading files (Attachment Harvester)

### Basic Usage

1. Run `PS1SNOWUtilities.ps1` (PowerShell 5.1 / STA recommended).
2. In the **Settings** tab, configure:
   - ServiceNow instance name
   - Authentication method (User ID + Password or API Key)
   - UI language if needed
3. Inputs are auto-saved to `settings.json` generated after first run (the file is not tracked in this repository).

##### Custom domain setup (when not using `instance-name.service-now.com`)

Add `instanceDomain` to `settings.json` to explicitly control the API base URL.

```json
{
  "instanceName": "dev12345",
  "instanceDomain": "example.com"
}
```

Or leave `instanceName` empty and provide a full URL with `https://` in `instanceDomain`.

```json
{
  "instanceName": "",
  "instanceDomain": "https://example.com"
}
```

- When `instanceDomain` is set, it takes precedence.
- You can set `instanceDomain` as either `example.com` or `https://example.com` (`https://` is automatically added if omitted).
- When `instanceDomain` is missing or empty, the app keeps the previous behavior and builds `https://<instanceName>.service-now.com` from `instanceName`.

#### Export workflow

1. In the **Export** tab, select the target table (or type it manually).
2. Optionally set filters (All records or `sys_updated_on` date range).
3. Choose an export directory and output format (CSV / JSON / Excel), then click **Execute**.
4. Check logs and use **Open Folder** to view exported files.

##### How to use split CSV export

1. Select **CSV** as the output format.
2. Enable **Split CSV Export** and set the number of rows per file.
3. Run export to generate multiple numbered CSV files in sequence.
4. Check logs for each generated file, then merge/process them as needed.

> 💡 Typical use case: when exporting a huge table to a single file may get cut off due to network or processing limits, split CSV export helps you safely output the full dataset in chunks.

#### Attachment Harvester workflow

1. In the **Attachment Harvester** tab, select a target table (or type it manually).
2. Choose the date field used for filtering (for example, `sys_updated_on`) and set start/end timestamps.
3. Select a download directory and optionally enable **Create subfolder per table**.
4. Run the harvester to download attachments linked to matched records; duplicate content is skipped using hash comparison.
5. Review logs for saved/skipped/failed counts.

> 💡 Typical use case: collect evidence files for incident review or audit requests across records updated during a defined period.

#### Database View Editor workflow

1. In the **Database View Editor** tab, enter the View name and label.
2. Select a base table, and set the base prefix if required.
3. Use **Add Join** to define join table, left/right columns, variable prefix, and LEFT JOIN options.
4. Click **Reload Columns** to refresh column candidates (currently, the loaded candidates are treated as visible columns as-is).
5. Click **Create View**, then review completion logs and links (created View list / View definition record).

#### Truncate (Delete all) workflow

1. In the **Truncate (Delete all)** tab, select the target table (or type it manually).
2. Check the read-only allowed-instance setting in the UI. To edit it, modify `truncateAllowedInstances` directly in `settings.json` (wildcards and comma-separated multiple patterns are supported; default: `*dev*,*stg*`).
3. Set max retry count (1-999), then click **Execute Delete All Records**.
4. In the displayed verification-code dialog, enter the 4-character code, then approve the execution confirmation dialog.
5. Check the progress bar and logs, and review retry logs as needed (up to the max retry count).

> 💡 Typical use case: You want to repeatedly delete data after large-volume import tests (tens of thousands of records) in development, but record-by-record deletion from the table management screen is too time-consuming.

> ⚠️ This feature is not recommended for production environments. Use it only for limited scenarios such as repeated large-volume import tests in development environments.

#### How to exclude specific features for internal distribution

- If you want to distribute the tool without high-risk features (for example, Truncate), exclude the corresponding feature file under `modules/Features` (for example, `TruncateFeature.psm1`) from the distribution package.
- Excluded features are not loaded at startup, so the related tab/actions will not appear in the UI.
- This allows you to ship a reduced-function build from the same codebase based on the target audience and operational policy.

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
