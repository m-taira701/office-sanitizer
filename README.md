# office-sanitizer

Office成果物（主に Excel / PowerPoint）を**納品前に安全な状態へ整形・サニタイズ**するためのCLIツール。

- 作成者・編集履歴・会社名などのメタデータを除去
- Excelの表示状態（倍率・選択セル）を統一
- 人手作業を前提にしない（CI/CDで自動実行できる設計）

---

## Motivation / Why

Officeファイル（.xlsx / .pptx など）には、次のような情報が意図せず残りやすい。

- 作成者 / 最終更新者
- 作成日時 / 更新日時
- 会社名・組織名（app properties）
- 表示倍率・選択セル・表示位置
- コメント・ノート

納品前に毎回これらを**人手で確認・削除するのは再現性が低く、事故りやすい**。

このツールは、

> 「納品前チェックリスト」をコード化し、
> **誰が・どこで実行しても同じ状態**を作る

ことを目的としている。

---

## Current Status

### 対応状況（v0.x）

- 対応ファイル
  - ✅ Excel (.xlsx)
  - ❌ PowerPoint (.pptx)（未実装）
  - ❌ Word (.docx)（未実装）

### Excelで行っている処理

- 表示状態の正規化
  - 表示倍率を 100% に設定
  - アクティブセルを A1 に設定
  - 最初のシートをアクティブに設定

- コメント削除
  - セルコメントを削除
  - Threaded comment は best-effort

- メタデータ削除
  - `docProps/core.xml`（作成者・日時など）
  - `docProps/custom.xml`（カスタムプロパティ）
  - `docProps/app.xml`（Company / Manager 等）

※ openpyxl だけに依存せず、zip を直接再構築することで**確実性を優先**している。

---

## Non-Goals（やらないこと）

- Officeファイルの完全匿名化
- マクロ（VBA）の解析・改変
- 表示内容（セル値・数式・レイアウト）の変更
- 既存のドキュメント構造の最適化

---

## Usage

### インストール（予定）

```bash
pip install office-sanitizer
```

※ 現状はリポジトリを直接 clone しての利用を想定。

### ローカル実行

```bash
office-sanitizer ./docs
```

- 指定したディレクトリ配下の `.xlsx` を再帰的に処理
- `~$` で始まる Office 一時ファイルは自動的に除外

### 単一ファイル

```bash
office-sanitizer ./sample.xlsx
```

### 挙動

- デフォルトでは **in-place（上書き）**
- 内部的には temp ファイルを経由して安全に置換

---

## Planned Features / Next Steps

### 1. PowerPoint (.pptx) 対応

- コメント削除
- ノート削除
- メタデータ（docProps）削除
- 初期表示スライドの正規化

### 2. Word (.docx) 対応

- メタデータ削除
- コメント削除
- トラックバック（変更履歴）削除

### 3. CI/CD 統合

- GitHub Actions
- GitLab CI / Runner

例：

```yaml
- uses: yourname/office-sanitizer@v1
  with:
    path: docs/
```

### 4. 実行モード拡張

- dry-run（変更内容の確認のみ）
- git diff 対象ファイルのみ処理

---

## Design Notes

- `.xlsx / .pptx / .docx` は zip 形式
- ライブラリレベルのAPIだけでは消しきれないメタデータがある
- そのため、本ツールでは
  - ライブラリ（openpyxl 等）
  - zip 直編集（docProps 再構築）

を併用している。

---

## License

MIT License
