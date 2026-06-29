# Windows Local Word Add-in

Windows 版 Word 専用の完全ローカル `.dotm` アドインです。

## インストール

`install_word_pdf_copy_addin.bat` を右クリックして `管理者として実行` ではなく、通常実行してください。

インストール後、Word を再起動します。リボンに `PDF作成` タブが表示されます。

## アンインストール

`uninstall_word_pdf_copy_addin.bat` を実行してください。

## 注意

このインストーラーは Word の VBA テンプレートアドインを自動生成します。生成時だけ Word の VBA プロジェクトへコードを挿入する必要があるため、現在のユーザーの Word 設定に `AccessVBOM` を一時的に有効化します。

作成されるファイル:

```text
%APPDATA%\Microsoft\Word\STARTUP\WordPdfCopyAddin.dotm
```

処理結果の PDF / DOCX は以下にも保存されます。

```text
%USERPROFILE%\Documents\WordPdfCopyExports
```

ボタン実行後、エクスプローラーやメール添付欄などに貼り付けると PDF と DOCX がファイルとして貼り付けられます。
