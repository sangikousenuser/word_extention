# Word PDF Copy Add-in

Word デスクトップ版のリボンに `PDF作成` タブを追加し、現在の文書を保存して PDF と DOCX を作成する Office アドインです。

## 開発実行

```bash
npm install
npm run dev
```

別ターミナルで Word デスクトップ版にサイドロードします。

```bash
npm run start:desktop
```

停止するとき:

```bash
npm run stop:desktop
```

## サーバーを毎回起動しない使い方

Office アドインは HTML / JavaScript / CSS を HTTPS URL から読み込む仕組みです。ローカル開発サーバーを起動せずに使う場合は、`dist` を GitHub Pages、Cloudflare Pages、Azure Static Web Apps などの静的ホスティングへ配置し、その公開 URL を指す manifest を Word に登録します。

GitHub Pages に出す場合は、GitHub の repository Settings > Pages で Source を `GitHub Actions` にしてください。このリポジトリには `.github/workflows/deploy-pages.yml` が入っているため、`main` ブランチへ push すると自動で `dist` が Pages に公開されます。

```bash
npm run build
ADDIN_BASE_URL=https://your-domain.example npm run build:production-manifest
```

GitHub Pages の URL が `https://your-name.github.io/word_extention/` の場合は、以下のように manifest を作ります。

```bash
ADDIN_BASE_URL=https://your-name.github.io/word_extention npm run build:production-manifest
```

この場合、Word に登録するファイルは `manifest.production.xml` です。以後は GitHub Pages 上のファイルを Word が読み込むため、手元で `npm run dev` を起動する必要はありません。

完全にローカル単体で、Web サーバーも HTTPS ホスティングも使わない形式にしたい場合は、Office.js アドインではなく Windows 専用の VSTO / COM アドインとして作る必要があります。VSTO / COM なら OS のファイル操作やクリップボード操作も強く扱えますが、macOS 版 Word では動きません。

## 動作

1. 現在の文書を保存します。
2. Word から PDF を取得します。
3. Word から DOCX を取得します。
4. Clipboard API が許可される環境では、PDF と DOCX の Blob をクリップボードへ書き込みます。
5. クリップボードへのファイルコピーが許可されない環境では、PDF と DOCX をダウンロードします。

Office Web アドインは OS ネイティブの Finder/Explorer ファイルコピーを直接操作できないため、クリップボード処理はホスト環境の Clipboard API 対応状況に依存します。

リボンの `PDF作成` ボタンは作業ウィンドウを開き、処理を自動実行します。作業ウィンドウ内のボタンから手動で再実行することもできます。
