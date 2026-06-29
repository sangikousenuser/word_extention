import { createPdfAndDocx, getErrorMessage } from "./exportFlow.js";

Office.onReady(() => {
  Office.actions.associate("createPdfAndCopy", async (event) => {
    try {
      await createPdfAndDocx();
    } catch (error) {
      console.error(error);
      window.alert(`PDF作成コピーに失敗しました: ${getErrorMessage(error)}`);
    } finally {
      event.completed();
    }
  });
});
