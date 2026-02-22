# 新UI改善ロードマップ（Codex実行用）

このドキュメントは、`PS1SNOWUtilities.ps1` ベースのWinForms UIを、
**段階的かつ安全に** 改善するための「Codexにそのまま渡せる実行手順」です。

---

## 0. ゴール

- UIを「固定座標中心」から「レイアウトコンテナ中心」に移行する。
- 見た目（おしゃれさ）を、テーマ・余白・情報階層の統一で底上げする。
- 大規模改修でも機能回帰を抑える。

---

## 1. 実行ルール（Codex向け）

1. 1フェーズごとにコミットする。
2. 変更対象は原則 `PS1SNOWUtilities.ps1` と `docs/` に限定する。
3. 既存機能（Export / Attachment / View Editor / Truncate / Logs / Settings）の文言キーは壊さない。
4. 各フェーズで最低1つの自己検証コマンドを実行する。
5. UI挙動が変わる場合、変更点を箇条書きでPR本文に明記する。

---

## 2. フェーズ別実装手順

## Phase 1: デザイン基盤の統一（低リスク）

### 目的
- テーマ・余白・ボタン種類を統一して、全体の見た目品質を上げる。

### 実装手順（Codexにそのまま指示可）
1. `Set-Theme` のパレットを整理し、`Back/Surface/Text/Muted/Accent/Danger/Border` の用途コメントを追加。
2. `Apply-ThemeRecursive` で以下を統一適用。
   - Label系は `Muted` を使う補助ラベル用ヘルパーを追加。
   - Button系は既定を Secondary 見た目に固定。
3. `Set-ButtonStyle` の呼び出しを全主要ボタンへ適用し、Primary/Danger/Secondary を明示。
4. フォントサイズを「通常9、見出し10太字」で統一。

### 完了条件
- 主要実行ボタンがPrimary、危険操作がDanger、それ以外がSecondaryになっている。
- Dark/Light切替時に文字色コントラストが破綻しない。

### 検証コマンド例
- `pwsh -NoProfile -Command "[void][scriptblock]::Create((Get-Content -Raw PS1SNOWUtilities.ps1))"`
- `rg -n "Set-ButtonStyle|Apply-ThemeRecursive|Set-Theme" PS1SNOWUtilities.ps1`

---

## Phase 2: レイアウト再構成（中リスク）

### 目的
- 固定座標依存を減らし、リサイズに強いUIへ移行。

### 実装手順
1. Exportタブから着手し、`Panel + Location/Size` 中心を `TableLayoutPanel` ベースに置換。
2. 入力エリアを以下3ブロックに分離。
   - 上段: 必須入力（テーブル、期間、出力先）
   - 中段: オプション（最大行数、形式、BOM等）
   - 下段: 実行/フォルダを開く
3. 各ブロックに共通余白（8/12/16）ルールを適用。
4. Anchor依存を減らし、Dock中心で再配置。

### 完了条件
- フォーム横幅を最小〜最大に変更しても、入力欄が重ならない。
- Exportタブの主要操作がスクロールなしで判読できる。

### 検証コマンド例
- `rg -n "TableLayoutPanel|FlowLayoutPanel" PS1SNOWUtilities.ps1`
- `rg -n "Location = New-Object System.Drawing.Point\(" PS1SNOWUtilities.ps1`

---

## Phase 3: 情報設計の統一（中リスク）

### 目的
- 全タブで「迷わない操作導線」に統一。

### 実装手順
1. 各タブに「セクション見出し」を導入（例: 入力 / オプション / 実行結果）。
2. 補助説明文を短文化し、長文はTooltipへ退避。
3. 主要アクションを右側 or 下部に固定し、タブ間で位置を揃える。
4. Logsタブの操作ボタン（コピー/クリア）を視認しやすい位置へ整理。

### 完了条件
- タブごとに操作構造が同型（入力→実行→確認）になっている。
- 説明文が過密にならず、初見でも操作順がわかる。

### 検証コマンド例
- `rg -n "New-UiSectionTitle|ToolTip|LogCopy|LogClear" PS1SNOWUtilities.ps1`

---

## Phase 4: 保守性強化（将来拡張）

### 目的
- UI改善を継続できる構成へ寄せる。

### 実装手順
1. UI構築処理を関数分割（例: `Build-ExportTab`, `Build-SettingsTab`）。
2. テーマ関連処理を1セクションに集約し、重複処理を削除。
3. タブごとの初期化処理を分離し、イベント登録点を明確化。
4. `docs/` に「UI変更時のチェックリスト」を追記。

### 完了条件
- `PS1SNOWUtilities.ps1` 内で「画面構築」「イベント登録」「ビジネス実行」の境界が読める。
- 次のUI改修時に変更点の局所化ができる。

### 検証コマンド例
- `rg -n "function Build-|add_Click|Apply-ThemeRecursive" PS1SNOWUtilities.ps1`

---

## 3. PRテンプレート（Codexが使う想定）

### タイトル例
- `UI: unify theme tokens and refactor Export layout to container-based design`

### 本文テンプレート
1. 何を変えたか（箇条書き3〜6件）
2. なぜ変えたか（可読性/操作性/保守性）
3. 回帰影響（既存機能への影響有無）
4. 実行した検証コマンドと結果

---

## 4. 変更時チェックリスト

- [ ] Dark/Light の双方で可読性OK
- [ ] Primary/Danger/Secondaryボタンの意味が統一
- [ ] ウィンドウリサイズ時に崩れない
- [ ] 文言キー（i18n）を壊していない
- [ ] 主要導線（入力→実行→確認）が維持されている

---

## 5. 非目標（今回やらない）

- 機能追加（新規タブ・新規API仕様追加）
- 認証仕様変更
- ServiceNow側仕様に依存するロジック変更

上記はUI改善と分離し、別PRで実施すること。
