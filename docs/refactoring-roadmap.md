# 将来の複数画面化を見据えたリファクタリング観点（PS1 SNOW Utilities）

## 前提
- 現状は `PS1SNOWUtilities.ps1` 単一ファイル（約 2,750 行）に、
  - WinForms 画面構築
  - ServiceNow API 呼び出し
  - 設定永続化
  - i18n
  - ドメインロジック（Export / View Editor / Truncate）
  が同居している。
- 画面追加（複数画面化）時に、**影響範囲が広くなる構造**になっているため、機能分割と依存関係整理を優先する。

---

## 優先度A（まず着手）

### 1) ファイル分割（レイヤ分離）
**課題**
- UI・API・設定・業務ロジックが1ファイルに混在し、変更時に副作用を読み切りにくい。

**提案**
- `modules/` を作り、最低限以下へ分離する。
  - `modules/Core/Settings.psm1`（Load/Save, DPAPI）
  - `modules/Core/ServiceNowClient.psm1`（GET/POST/PATCH/DELETE 共通）
  - `modules/Core/Logging.psm1`
  - `modules/Features/ExportFeature.psm1`
  - `modules/Features/ViewEditorFeature.psm1`
  - `modules/Features/TruncateFeature.psm1`
  - `modules/UI/MainForm.psm1`（フォーム組み立て）

**効果**
- 画面追加時に Feature 単位で独立実装しやすくなる。
- テスト可能領域（UI非依存）が増える。

---

### 2) グローバル状態（`$script:`）依存の縮小
**課題**
- `$script:Settings`, `$script:ColumnCache`, `$script:txtLog` などの共有状態が多く、イベント追加時に状態競合を起こしやすい。

**提案**
- `AppContext`（単一の状態オブジェクト）を導入し、依存を明示して関数へ渡す。
- APIクライアントや設定サービスはコンストラクタ相当の初期化関数で生成。

**効果**
- 依存が可視化され、機能追加時の壊れやすさを減らせる。

---

### 3) ServiceNow API呼び出しの共通化
**課題**
- 認証処理と `Invoke-RestMethod` の重複が多い（GET/POST/PATCH/DELETEで同型コード）。

**提案**
- `Invoke-SnowRequest -Method -Path -Body -TimeoutSec` に統合。
- 認証（userpass / apikey）を戦略化して切り替える。
- エラーオブジェクトを標準化（`Code`, `Message`, `Operation`, `Hint`）。

**効果**
- 新画面で API を増やしても、実装コストとバグ混入率を抑制できる。

---

### 4) 画面（View）とユースケース（ロジック）の分離
**課題**
- ボタンクリック内でデータ取得・検証・変換・描画更新が混ざるため、再利用しにくい。

**提案**
- 各機能を `UseCase` 関数化（例: `Invoke-ExportUseCase`, `Invoke-CreateViewUseCase`）。
- UI層は入力収集と結果反映のみ担当。

**効果**
- 同一ロジックを将来の別画面（ウィザード/詳細画面）へ再利用可能。

---

## 優先度B（中期）

### 5) i18n辞書の外部化
**課題**
- 文言辞書が本体コード内にあり、文言修正の差分が大きくレビューしづらい。

**提案**
- `locales/ja.json`, `locales/en.json` へ分離。
- `Get-Text key` で取得し、未定義キー検出ログを追加。

**効果**
- 多言語対応・文言修正を安全に運用できる。

---

### 6) 設定モデルのバージョニング
**課題**
- `settings.json` の項目追加時に後方互換ロジックが暗黙的。

**提案**
- `settingsVersion` を導入し、`Migrate-Settings` を実装。
- マイグレーション単位で関数分割（v1→v2, v2→v3）。

**効果**
- 将来機能追加時の設定破損リスクを低減。

---

### 7) バリデーション統一
**課題**
- 入力検証が機能ごとに散在し、エラー文言・判定粒度が不揃い。

**提案**
- `Validate-ExportInput`, `Validate-ViewInput`, `Validate-TruncateInput` へ統一。
- 返却型を `ValidationResult`（`IsValid`, `Errors[]`）に揃える。

**効果**
- 画面追加時に同じ検証フレームを使い回せる。

---

### 8) ログの構造化
**課題**
- 現状はプレーン文字列中心で、障害調査時のフィルタが難しい。

**提案**
- UIログに加え、ファイルへ JSON Lines 形式出力（任意ON/OFF）。
- `level`, `feature`, `operation`, `durationMs`, `errorCode` を標準項目化。

**効果**
- 問題切り分け速度が向上し、保守性が上がる。

---

## 優先度C（余力があれば）

### 9) 非同期実行基盤の見直し
**課題**
- `Invoke-Async` が実質同期実行に近く、処理時間が伸びるとUI応答性に影響。

**提案**
- `BackgroundWorker` あるいは Runspace を採用し、
  進捗通知・キャンセル制御を標準化。

### 10) プラグイン的な画面追加方式
**提案**
- `features/*.psm1` を読み込んでタブを動的登録する方式へ。
- 画面追加は `Register-FeatureTab` 実装のみで完結させる。

### 11) テスト基盤（Pester）
**提案**
- まずはUI非依存層（Settings, URL/Query生成, Validation, Mapping）からテスト導入。

---

## 推奨実行順（現実的な3フェーズ）
1. **Phase 1（安全分割）**
   - APIクライアント共通化
   - Settings/Logging 分離
   - 既存挙動維持を最優先（UIは極力触らない）
2. **Phase 2（機能単位分割）**
   - Export / View Editor / Truncate を UseCase 化
   - バリデーション統一
3. **Phase 3（拡張運用対応）**
   - i18n外部化
   - 設定マイグレーション
   - テスト追加

---

## 追加で先に決めると良い設計方針
- 画面追加の単位：タブ追加か、別フォーム起動か。
- 共通状態管理：`AppContext` を唯一の共有コンテナにするか。
- APIエラーハンドリング：ユーザ向けメッセージと詳細ログの分離方針。
- 互換性方針：PowerShell 5.1 固定を維持するか（7系対応を視野に入れるか）。

