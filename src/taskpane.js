import "./taskpane.css";
import { createPdfAndDocx, getErrorMessage } from "./exportFlow.js";

const statusElement = document.querySelector("#status");
const runButton = document.querySelector("#runButton");

Office.onReady(() => {
  setStatus("準備できました。");
  runButton.disabled = false;
  runButton.addEventListener("click", () => runExportFlow());
});

async function runExportFlow() {
  runButton.disabled = true;

  try {
    const result = await createPdfAndDocx({ onStatus: setStatus });
    const names = result.files.map((file) => file.name).join(" と ");

    if (result.copied) {
      setStatus(`完了しました。${names} をクリップボードへコピーしました。`);
      return;
    }

    setStatus("クリップボードへのファイルコピーが許可されなかったため、PDF と DOCX をダウンロードしました。");
  } catch (error) {
    console.error(error);
    setStatus(`失敗しました: ${getErrorMessage(error)}`);
  } finally {
    runButton.disabled = false;
  }
}

function setStatus(message) {
  statusElement.textContent = message;
}
