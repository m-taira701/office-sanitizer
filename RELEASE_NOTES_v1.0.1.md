# v1.0.1 - 誤検知対策 & 安定性向上

v1.0.0 のマイナーアップデートです。Windows Defender等のセキュリティソフトによる誤検知（False Positive）を軽減し、より安定してアプリを提供できるようにするため、ビルドツールを PyInstaller に変更しました。

※ 起動速度は v1.0.0 より若干遅くなる可能性がありますが、その他の機能は同一です。

## 🛡️ 変更点 (Changes)

*   **ビルドツールの変更 (PyInstaller)**: アプリケーションのパッケージングツールを広く使われているPyInstallerへ変更しました。これにより一部のセキュリティソフトでの誤検知（`Trojan:Win32/Bearfoos.B!ml` 等）が低減されることが期待されます。
    *   *注意: コード署名証明書による署名は行っていないため、依然としてSmartScreen警告が出る場合がありますが、安全に実行できます。*

## 🚀 v1.0.0 の主要機能 (Recap)

*   **モダンなGUI**: PySide6ベースのダークテーマGUIを搭載。ドラッグ＆ドロップ対応。
*   **安定した実行**: 煩雑な環境構築なしに利用できる単一実行ファイル。
*   **設定パネル**: メタデータ、コメント、変更履歴の削除有無を個別に設定可能。

## 📦 対応フォーマット

*   Microsoft Word (.docx)
*   Microsoft Excel (.xlsx)
*   Microsoft PowerPoint (.pptx)

## 📥 インストール

下の Assets から、お使いのOSに合わせたファイルをダウンロードしてください。
*   **Windows**: `OfficeSanitizer-Windows.exe`
*   **macOS**: `OfficeSanitizer-Mac`
