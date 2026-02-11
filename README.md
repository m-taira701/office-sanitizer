# Office Sanitizer

Officeファイル（Excel, Word, PowerPoint）のメタデータ、コメント、変更履歴を一括削除するデスクトップアプリケーション。

## 特徴
- **堅牢なサニタイズ処理**: 内部構造（XML）を直接操作し、メタデータやコメントを確実に削除。
- **モダンなGUI (PySide6)**: ダークモード対応、ドラッグ&ドロップ対応。
- **CLI対応**: 自動化スクリプト等で使えるコマンドラインインターフェース。
- **クロスプラットフォーム**: Windows/Mac対応（Pythonが動く環境）。

## 必要要件
- Python 3.8+ (推奨: **Python 3.13**)
- PySide6
- openpyxl

## インストール

```bash
pip install -r requirements.txt
```

## 使い方

### GUIモード

```bash
python main.py
```
起動後、ファイルをドラッグ&ドロップし、設定を確認して「サニタイズ実行」ボタンを押してください。

### CLIモード

```bash
# ヘルプ表示
python main.py --help

# ファイル単体処理
python main.py target.docx

# ディレクトリ一括処理（再帰的）
python main.py ./docs_folder -r

# 元ファイルを上書きする場合
python main.py target.xlsx --inplace
```

## ビルド (exe/app化)

[Nuitka](https://nuitka.net/) を使用して高速なシングルバイナリを作成します。
**注意**: 安定動作のため、**Python 3.13** 環境でのビルドを推奨します。

1. アイコンを設定する場合: `resources/icon.ico` (Windows) または `resources/icon.icns` (Mac) を配置してください。
2. ビルドスクリプトを実行:

```bash
# Nuitkaとzstandardがインストールされていることを確認
pip install nuitka zstandard

python build.py
```

`dist/` フォルダに実行ファイル (`OfficeSanitizer.exe` または `OfficeSanitizer.bin`) が生成されます。
初回ビルド時はCコンパイラのダウンロードが走る場合があります。

## プロジェクト構成

- `main.py`: エントリーポイント
- `build.py`: Nuitkaビルドスクリプト
- `office_sanitizer/core/`: サニタイズロジック（独立モジュール）
- `office_sanitizer/gui/`: PySide6 GUI実装
- `office_sanitizer/cli.py`: CLI実装
- `resources/`: アイコン等のリソースファイル置き場
