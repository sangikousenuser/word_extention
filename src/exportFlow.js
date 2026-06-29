export async function createPdfAndDocx(options = {}) {
  const { onStatus = () => {} } = options;

  onStatus("文書を保存しています...");
  await saveCurrentDocument();

  onStatus("PDF を作成しています...");
  const pdfBlob = await getDocumentBlob(Office.FileType.Pdf, "application/pdf");

  onStatus("DOCX を作成しています...");
  const docxBlob = await getDocumentBlob(
    Office.FileType.Compressed,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  );

  const baseName = getDocumentBaseName();
  const files = [
    {
      name: `${baseName}.pdf`,
      blob: pdfBlob,
      mimeType: "application/pdf",
    },
    {
      name: `${baseName}.docx`,
      blob: docxBlob,
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    },
  ];

  onStatus("クリップボードへコピーしています...");
  const copied = await copyFilesToClipboard(files);

  if (!copied) {
    files.forEach((file) => downloadBlob(file.blob, file.name));
  }

  return { copied, files };
}

export function getErrorMessage(error) {
  if (!error) {
    return "不明なエラー";
  }

  return error.message || error.name || String(error);
}

function saveCurrentDocument() {
  return new Promise((resolve, reject) => {
    Office.context.document.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(result.error);
    });
  });
}

function getDocumentBlob(fileType, mimeType) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, { sliceSize: 65536 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error);
        return;
      }

      readFileSlices(result.value, mimeType).then(resolve, reject);
    });
  });
}

async function readFileSlices(file, mimeType) {
  const chunks = [];

  try {
    for (let index = 0; index < file.sliceCount; index += 1) {
      const slice = await getSlice(file, index);
      chunks.push(new Uint8Array(slice.data));
    }

    return new Blob(chunks, { type: mimeType });
  } finally {
    file.closeAsync();
  }
}

function getSlice(file, index) {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      reject(result.error);
    });
  });
}

async function copyFilesToClipboard(files) {
  if (!navigator.clipboard || typeof ClipboardItem === "undefined") {
    return false;
  }

  try {
    const itemData = Object.fromEntries(files.map((file) => [file.mimeType, file.blob]));
    await navigator.clipboard.write([new ClipboardItem(itemData)]);
    return true;
  } catch (error) {
    console.warn("Clipboard write failed; falling back to downloads.", error);
    return false;
  }
}

function getDocumentBaseName() {
  const url = Office.context.document.url || "";
  const fallback = "document";
  const lastPart = decodeURIComponent(url.split(/[\\/]/).pop() || fallback);
  return sanitizeFileName(lastPart.replace(/\.[^.]+$/, "") || fallback);
}

function sanitizeFileName(name) {
  return name.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").trim() || "document";
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
