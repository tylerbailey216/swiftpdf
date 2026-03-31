(function () {
  "use strict";

  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const pdfjs = window.pdfjsLib;

  if (pdfjs && pdfjs.GlobalWorkerOptions) {
    pdfjs.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.js";
  }

  const LETTER_PORTRAIT = { width: 612, height: 792 };
  const LETTER_LANDSCAPE = { width: 792, height: 612 };
  const DOCX_LAYOUT = {
    width: LETTER_PORTRAIT.width,
    height: LETTER_PORTRAIT.height,
    marginX: 54,
    marginY: 54,
    fontSize: 11,
    lineHeight: 16,
  };

  const fileInput = document.getElementById("fileInput");
  const dropzone = document.getElementById("dropzone");
  const browseBtn = document.getElementById("browseBtn");
  const addBlankPageBtn = document.getElementById("addBlankPageBtn");
  const removePageBtn = document.getElementById("removePageBtn");
  const pageList = document.getElementById("pageList");
  const pageEmpty = document.getElementById("pageEmpty");
  const statusEl = document.getElementById("status");
  const editorEmpty = document.getElementById("editorEmpty");
  const editorStageWrap = document.getElementById("editorStageWrap");
  const pageFrame = document.getElementById("pageFrame");
  const pdfCanvas = document.getElementById("pdfCanvas");
  const imagePreview = document.getElementById("imagePreview");
  const docxPreview = document.getElementById("docxPreview");
  const blankPreview = document.getElementById("blankPreview");
  const fieldLayer = document.getElementById("fieldLayer");
  const selectedPageTitle = document.getElementById("selectedPageTitle");
  const selectedPageMeta = document.getElementById("selectedPageMeta");
  const addTextFieldBtn = document.getElementById("addTextFieldBtn");
  const addTextareaFieldBtn = document.getElementById("addTextareaFieldBtn");
  const addCheckboxFieldBtn = document.getElementById("addCheckboxFieldBtn");
  const addSignatureFieldBtn = document.getElementById("addSignatureFieldBtn");
  const fieldEmpty = document.getElementById("fieldEmpty");
  const fieldSettings = document.getElementById("fieldSettings");
  const fieldLabelInput = document.getElementById("fieldLabelInput");
  const fieldSummary = document.getElementById("fieldSummary");
  const removeFieldBtn = document.getElementById("removeFieldBtn");
  const signatureNameInput = document.getElementById("signatureNameInput");
  const signatureUploadInput = document.getElementById("signatureUploadInput");
  const signaturePreviewCard = document.getElementById("signaturePreviewCard");
  const signaturePreviewCopy = document.getElementById("signaturePreviewCopy");
  const signaturePreviewImage = document.getElementById("signaturePreviewImage");
  const clearSignatureBtn = document.getElementById("clearSignatureBtn");
  const buildBtn = document.getElementById("buildBtn");
  const downloadLink = document.getElementById("downloadLink");

  const measureCanvas = document.createElement("canvas");
  const measureContext = measureCanvas.getContext("2d");

  const state = {
    documents: [],
    pages: [],
    nextDocumentId: 1,
    nextPageId: 1,
    nextFieldId: 1,
    selectedPageId: null,
    selectedFieldId: null,
    downloadUrl: null,
    previewToken: 0,
    pdfPreviewCache: new Map(),
    interaction: null,
    signatureName: "",
    signatureDataUrl: null,
    signatureMimeType: "",
    signatureBytes: null,
  };

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function setStatus(message) {
    if (statusEl) {
      statusEl.textContent = message;
    }
  }

  function slugify(value) {
    return (
      String(value || "")
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "") || "field"
    );
  }

  function formatKindLabel(kind) {
    if (kind === "pdf") return "PDF";
    if (kind === "image") return "Image";
    if (kind === "docx") return "Word";
    return "Blank";
  }

  function formatFieldType(kind) {
    if (kind === "text") return "Short answer";
    if (kind === "textarea") return "Paragraph box";
    if (kind === "signature") return "Signature";
    return "Checkbox";
  }

  function formatPageSize(width, height) {
    const widthIn = (width / 72).toFixed(1);
    const heightIn = (height / 72).toFixed(1);
    return widthIn + " x " + heightIn + " in";
  }

  function dataUrlToBytes(dataUrl) {
    if (!dataUrl || typeof dataUrl !== "string" || dataUrl.indexOf(",") < 0) {
      return null;
    }
    const payload = dataUrl.split(",")[1];
    if (!payload) {
      return null;
    }
    const binary = atob(payload);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  }

  function resetDownloadLink() {
    if (state.downloadUrl) {
      URL.revokeObjectURL(state.downloadUrl);
      state.downloadUrl = null;
    }
    downloadLink.href = "#";
    downloadLink.classList.add("disabled");
  }

  function setDownloadLink(blobUrl, fileName) {
    resetDownloadLink();
    state.downloadUrl = blobUrl;
    downloadLink.href = blobUrl;
    downloadLink.download = fileName;
    downloadLink.classList.remove("disabled");
  }

  function getSelectedPage() {
    return (
      state.pages.find(function (page) {
        return page.id === state.selectedPageId;
      }) || null
    );
  }

  function getSelectedField() {
    const page = getSelectedPage();
    if (!page) return null;
    return (
      page.fields.find(function (field) {
        return field.id === state.selectedFieldId;
      }) || null
    );
  }

  function getDocumentById(documentId) {
    return (
      state.documents.find(function (document) {
        return document.id === documentId;
      }) || null
    );
  }

  function hideAllBackgrounds() {
    pdfCanvas.hidden = true;
    imagePreview.hidden = true;
    docxPreview.hidden = true;
    blankPreview.hidden = true;
  }

  function setToolAvailability() {
    const hasPage = Boolean(getSelectedPage());
    const hasField = Boolean(getSelectedField());

    addTextFieldBtn.disabled = !hasPage;
    addTextareaFieldBtn.disabled = !hasPage;
    addCheckboxFieldBtn.disabled = !hasPage;
    addSignatureFieldBtn.disabled = !hasPage;
    removePageBtn.disabled = !hasPage;
    buildBtn.disabled = state.pages.length === 0;
    removeFieldBtn.disabled = !hasField;
  }

  function renderSignaturePreview() {
    const hasImage = Boolean(state.signatureDataUrl);
    signaturePreviewImage.hidden = !hasImage;
    if (hasImage) {
      signaturePreviewImage.src = state.signatureDataUrl;
      signaturePreviewCopy.textContent = state.signatureName
        ? "Signature image loaded. The exported PDF will place this image with the signer name " +
          state.signatureName +
          "."
        : "Signature image loaded. The exported PDF will place this image inside each signature field.";
      return;
    }

    signaturePreviewImage.removeAttribute("src");
    signaturePreviewCopy.textContent = state.signatureName
      ? "No signature image loaded. The exported PDF will use the typed signer name " +
        state.signatureName +
        " with a signature line."
      : "No signature image loaded yet. You can still use a typed signer name.";
  }

  function getDefaultImagePageSize(width, height) {
    return width >= height
      ? { width: LETTER_LANDSCAPE.width, height: LETTER_LANDSCAPE.height }
      : { width: LETTER_PORTRAIT.width, height: LETTER_PORTRAIT.height };
  }

  function fitIntoPage(contentWidth, contentHeight, pageWidth, pageHeight, margin) {
    const safeWidth = pageWidth - margin * 2;
    const safeHeight = pageHeight - margin * 2;
    const scale = Math.min(safeWidth / contentWidth, safeHeight / contentHeight);
    const width = contentWidth * scale;
    const height = contentHeight * scale;
    return {
      x: (pageWidth - width) / 2,
      y: (pageHeight - height) / 2,
      width: width,
      height: height,
    };
  }

  function makeDocument(record) {
    state.documents.push(record);
    return record;
  }

  function makePage(overrides) {
    return Object.assign(
      {
        id: state.nextPageId++,
        kind: "blank",
        sourceName: "Blank page",
        sourceDetail: "Blank page",
        width: LETTER_PORTRAIT.width,
        height: LETTER_PORTRAIT.height,
        fields: [],
      },
      overrides || {}
    );
  }

  function createBlankPage() {
    return makePage({
      kind: "blank",
      sourceName: "Blank page",
      sourceDetail: "Blank page ready for new form fields",
    });
  }

  function ensureSelection() {
    if (!state.pages.length) {
      state.selectedPageId = null;
      state.selectedFieldId = null;
      return;
    }

    if (
      !state.pages.some(function (page) {
        return page.id === state.selectedPageId;
      })
    ) {
      state.selectedPageId = state.pages[0].id;
    }

    const page = getSelectedPage();
    if (
      !page ||
      !page.fields.some(function (field) {
        return field.id === state.selectedFieldId;
      })
    ) {
      state.selectedFieldId = null;
    }
  }

  function selectPage(pageId) {
    state.selectedPageId = pageId;
    const page = getSelectedPage();
    if (
      !page ||
      !page.fields.some(function (field) {
        return field.id === state.selectedFieldId;
      })
    ) {
      state.selectedFieldId = null;
    }
    render();
  }

  function selectField(fieldId) {
    state.selectedFieldId = fieldId;
    renderFieldLayer();
    renderInspector();
    setToolAvailability();
  }

  function addPageRecords(pageRecords) {
    if (!pageRecords.length) {
      return;
    }
    pageRecords.forEach(function (pageRecord) {
      state.pages.push(pageRecord);
    });
    state.selectedPageId = pageRecords[0].id;
    state.selectedFieldId = null;
    resetDownloadLink();
    render();
  }

  function getFileKind(file) {
    const lowerName = String(file.name || "").toLowerCase();
    if (file.type === "application/pdf" || lowerName.endsWith(".pdf")) {
      return "pdf";
    }
    if (
      file.type === "image/jpeg" ||
      lowerName.endsWith(".jpg") ||
      lowerName.endsWith(".jpeg") ||
      lowerName.endsWith(".jfif") ||
      file.type === "image/png" ||
      lowerName.endsWith(".png")
    ) {
      return "image";
    }
    if (
      file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
      lowerName.endsWith(".docx")
    ) {
      return "docx";
    }
    if (file.type === "application/msword" || lowerName.endsWith(".doc")) {
      return "doc";
    }
    return "unsupported";
  }

  function loadImageFromBlob(blob) {
    return new Promise(function (resolve, reject) {
      const url = URL.createObjectURL(blob);
      const image = new Image();
      image.onload = function () {
        resolve({
          image: image,
          url: url,
          width: image.naturalWidth || image.width,
          height: image.naturalHeight || image.height,
        });
      };
      image.onerror = function () {
        URL.revokeObjectURL(url);
        reject(new Error("Image load failed"));
      };
      image.src = url;
    });
  }

  function wrapLineToWidth(text, maxWidth, fontDefinition) {
    const normalized = String(text || "").replace(/\s+/g, " ").trim();
    if (!normalized) {
      return [""];
    }

    measureContext.font = fontDefinition;

    const words = normalized.split(" ");
    const lines = [];
    let current = "";

    words.forEach(function (word) {
      const next = current ? current + " " + word : word;
      if (measureContext.measureText(next).width <= maxWidth) {
        current = next;
        return;
      }

      if (current) {
        lines.push(current);
      }

      if (measureContext.measureText(word).width <= maxWidth) {
        current = word;
        return;
      }

      let chunk = "";
      for (let index = 0; index < word.length; index += 1) {
        const candidate = chunk + word[index];
        if (measureContext.measureText(candidate).width <= maxWidth) {
          chunk = candidate;
        } else {
          if (chunk) {
            lines.push(chunk);
          }
          chunk = word[index];
        }
      }
      current = chunk;
    });

    if (current) {
      lines.push(current);
    }

    return lines.length ? lines : [""];
  }

  function createDocxLayouts(rawText) {
    const normalized = String(rawText || "")
      .replace(/\r/g, "")
      .replace(/\t/g, "    ")
      .replace(/\u00a0/g, " ")
      .trim();

    const paragraphSource = normalized ? normalized.split(/\n+/) : ["This page was imported from Word."];
    const wrappedLines = [];
    const fontDefinition = DOCX_LAYOUT.fontSize + "pt Helvetica, Arial, sans-serif";
    const maxWidth = DOCX_LAYOUT.width - DOCX_LAYOUT.marginX * 2;

    paragraphSource.forEach(function (paragraph) {
      const lines = wrapLineToWidth(paragraph, maxWidth, fontDefinition);
      lines.forEach(function (line) {
        wrappedLines.push(line);
      });
      wrappedLines.push("");
    });

    while (wrappedLines.length && wrappedLines[wrappedLines.length - 1] === "") {
      wrappedLines.pop();
    }

    const maxLinesPerPage = Math.max(
      1,
      Math.floor((DOCX_LAYOUT.height - DOCX_LAYOUT.marginY * 2) / DOCX_LAYOUT.lineHeight)
    );
    const pages = [];

    for (let index = 0; index < wrappedLines.length; index += maxLinesPerPage) {
      pages.push({
        lines: wrappedLines.slice(index, index + maxLinesPerPage),
        width: DOCX_LAYOUT.width,
        height: DOCX_LAYOUT.height,
      });
    }

    if (!pages.length) {
      pages.push({
        lines: ["This page was imported from Word."],
        width: DOCX_LAYOUT.width,
        height: DOCX_LAYOUT.height,
      });
    }

    return pages;
  }

  async function importPdf(file) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const pdfDocument = await PDFDocument.load(bytes);
    const documentRecord = makeDocument({
      id: state.nextDocumentId++,
      kind: "pdf",
      name: file.name,
      bytes: bytes,
    });

    const pages = [];
    for (let pageIndex = 0; pageIndex < pdfDocument.getPageCount(); pageIndex += 1) {
      const page = pdfDocument.getPage(pageIndex);
      const size = page.getSize();
      pages.push(
        makePage({
          kind: "pdf",
          documentId: documentRecord.id,
          sourceName: file.name,
          sourceDetail: "PDF page " + (pageIndex + 1),
          sourcePageIndex: pageIndex,
          width: size.width,
          height: size.height,
        })
      );
    }
    return pages;
  }

  async function importImage(file) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const blob = new Blob([bytes], { type: file.type || "image/jpeg" });
    const imageData = await loadImageFromBlob(blob);
    const pageSize = getDefaultImagePageSize(imageData.width, imageData.height);

    const documentRecord = makeDocument({
      id: state.nextDocumentId++,
      kind: "image",
      name: file.name,
      bytes: bytes,
      mimeType: file.type || "image/jpeg",
      objectUrl: imageData.url,
      imageWidth: imageData.width,
      imageHeight: imageData.height,
    });

    return [
      makePage({
        kind: "image",
        documentId: documentRecord.id,
        sourceName: file.name,
        sourceDetail: "Image page",
        width: pageSize.width,
        height: pageSize.height,
      }),
    ];
  }

  async function importDocx(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
    const layouts = createDocxLayouts(result && result.value ? result.value : "");
    const documentRecord = makeDocument({
      id: state.nextDocumentId++,
      kind: "docx",
      name: file.name,
    });

    return layouts.map(function (layout, index) {
      return makePage({
        kind: "docx",
        documentId: documentRecord.id,
        sourceName: file.name,
        sourceDetail: "Word page " + (index + 1),
        width: layout.width,
        height: layout.height,
        layout: layout,
      });
    });
  }

  async function addFiles(fileListLike) {
    const files = Array.prototype.slice.call(fileListLike || []);
    if (!files.length) {
      return;
    }

    let addedPages = [];
    let skippedLegacyDoc = 0;
    let skippedUnsupported = 0;

    for (const file of files) {
      const kind = getFileKind(file);
      try {
        if (kind === "pdf") {
          addedPages = addedPages.concat(await importPdf(file));
        } else if (kind === "image") {
          addedPages = addedPages.concat(await importImage(file));
        } else if (kind === "docx") {
          addedPages = addedPages.concat(await importDocx(file));
        } else if (kind === "doc") {
          skippedLegacyDoc += 1;
        } else {
          skippedUnsupported += 1;
        }
      } catch (error) {
        console.error(error);
        skippedUnsupported += 1;
      }
    }

    if (addedPages.length) {
      addPageRecords(addedPages);
    }

    if (addedPages.length && (skippedLegacyDoc || skippedUnsupported)) {
      setStatus(
        "Added " +
          addedPages.length +
          " page(s). Skipped " +
          skippedLegacyDoc +
          " legacy Word file(s) and " +
          skippedUnsupported +
          " unsupported file(s)."
      );
      return;
    }

    if (addedPages.length) {
      setStatus("Added " + addedPages.length + " page(s). Choose a page to start adding fields.");
      return;
    }

    if (skippedLegacyDoc) {
      setStatus("This browser tool needs Word files saved as .docx first.");
      return;
    }

    setStatus("Please add a PDF, JPG, PNG, or Word .docx file.");
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function renderPageList() {
    pageList.innerHTML = "";
    pageEmpty.hidden = state.pages.length > 0;

    state.pages.forEach(function (page, index) {
      const item = document.createElement("li");
      item.className = "page-card" + (page.id === state.selectedPageId ? " active" : "");
      item.setAttribute("data-page-id", String(page.id));

      const fieldCount = page.fields.length;
      const typeLabel = formatKindLabel(page.kind);

      item.innerHTML =
        '<button class="page-card-main" type="button" data-action="select-page" data-page-id="' +
        page.id +
        '">' +
        '<div class="page-card-top">' +
        '<span class="page-badge">' +
        typeLabel +
        "</span>" +
        '<span class="page-number">Page ' +
        (index + 1) +
        "</span>" +
        "</div>" +
        '<div class="page-title">' +
        escapeHtml(page.sourceName) +
        "</div>" +
        '<div class="page-copy">' +
        escapeHtml(page.sourceDetail) +
        " - " +
        fieldCount +
        " field" +
        (fieldCount === 1 ? "" : "s") +
        "</div>" +
        "</button>" +
        '<div class="page-actions">' +
        '<button class="ghost" type="button" data-action="move-up" data-page-id="' +
        page.id +
        '"' +
        (index === 0 ? " disabled" : "") +
        '>Move up</button>' +
        '<button class="ghost" type="button" data-action="move-down" data-page-id="' +
        page.id +
        '"' +
        (index === state.pages.length - 1 ? " disabled" : "") +
        '>Move down</button>' +
        '<button class="ghost" type="button" data-action="remove-page" data-page-id="' +
        page.id +
        '">Remove</button>' +
        "</div>";

      pageList.appendChild(item);
    });
  }

  function renderInspector() {
    const page = getSelectedPage();
    const field = getSelectedField();

    if (!page) {
      selectedPageTitle.textContent = "No page selected";
      selectedPageMeta.textContent = "Page details will appear here.";
      fieldEmpty.hidden = false;
      fieldSettings.hidden = true;
      fieldLabelInput.value = "";
      fieldSummary.textContent = "";
      return;
    }

    selectedPageTitle.textContent = page.sourceName;
    selectedPageMeta.textContent =
      page.sourceDetail +
      " - " +
      formatPageSize(page.width, page.height) +
      " - " +
      page.fields.length +
      " field" +
      (page.fields.length === 1 ? "" : "s");

    if (!field) {
      fieldEmpty.hidden = false;
      fieldSettings.hidden = true;
      fieldLabelInput.value = "";
      fieldSummary.textContent = "";
      return;
    }

    fieldEmpty.hidden = true;
    fieldSettings.hidden = false;
    fieldLabelInput.value = field.label;
    fieldSummary.textContent =
      formatFieldType(field.kind) +
      " on page " +
      (state.pages.findIndex(function (entry) {
        return entry.id === page.id;
      }) +
        1) +
      (field.kind === "signature"
        ? ". Drag it where the signature should appear. The export uses your uploaded signature image or typed signer name."
        : ". Drag it on the page to move it. Resize it from the bottom-right corner.");
  }

  function renderFieldLayer() {
    const page = getSelectedPage();
    fieldLayer.innerHTML = "";

    if (!page) {
      return;
    }

    page.fields.forEach(function (field) {
      const node = document.createElement("div");
      node.className = "field-box " + field.kind + (field.id === state.selectedFieldId ? " active" : "");
      node.setAttribute("data-field-id", String(field.id));
      node.style.left = field.x * 100 + "%";
      node.style.top = field.y * 100 + "%";
      node.style.width = field.width * 100 + "%";
      node.style.height = field.height * 100 + "%";
      let bodyMarkup = '<div class="field-body"></div>';
      if (field.kind === "signature") {
        bodyMarkup =
          '<div class="field-body"><div class="signature-stamp">' +
          (state.signatureDataUrl
            ? '<img src="' + state.signatureDataUrl + '" alt="" />'
            : '<div class="signature-stamp-name">' +
              escapeHtml(state.signatureName || "Signature") +
              "</div>") +
          "</div></div>";
      }
      node.innerHTML =
        '<span class="field-label">' +
        escapeHtml(field.label) +
        "</span>" +
        bodyMarkup +
        '<div class="field-resize" data-action="resize-field" aria-hidden="true"></div>';
      fieldLayer.appendChild(node);
    });
  }

  async function getPdfPreviewDocument(documentId) {
    if (state.pdfPreviewCache.has(documentId)) {
      return state.pdfPreviewCache.get(documentId);
    }

    const record = getDocumentById(documentId);
    if (!record || !pdfjs) {
      return null;
    }

    const loadingTask = pdfjs.getDocument({
      data: new Uint8Array(record.bytes),
      isEvalSupported: false,
      disableAutoFetch: true,
      disableStream: true,
      useWorkerFetch: false,
    });
    const promise = loadingTask.promise.catch(function (error) {
      state.pdfPreviewCache.delete(documentId);
      throw error;
    });
    state.pdfPreviewCache.set(documentId, promise);
    return promise;
  }

  async function renderPdfPageToCanvas(page) {
    const pdfDocument = await getPdfPreviewDocument(page.documentId);
    if (!pdfDocument) {
      throw new Error("PDF preview library is not available.");
    }

    const pdfPage = await pdfDocument.getPage(page.sourcePageIndex + 1);
    const frameWidth = pageFrame ? pageFrame.clientWidth : 0;
    const baseScale = frameWidth && page.width ? frameWidth / page.width : 1;
    const pixelRatio = window.devicePixelRatio || 1;
    const viewport = pdfPage.getViewport({
      scale: Math.max(1.1, baseScale * pixelRatio),
    });
    const context = pdfCanvas.getContext("2d", { alpha: false });

    pdfCanvas.width = Math.max(1, Math.floor(viewport.width));
    pdfCanvas.height = Math.max(1, Math.floor(viewport.height));
    pdfCanvas.style.width = "100%";
    pdfCanvas.style.height = "100%";
    context.save();
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, pdfCanvas.width, pdfCanvas.height);
    context.restore();

    await pdfPage.render({
      canvasContext: context,
      viewport: viewport,
      intent: "display",
    }).promise;
  }

  function renderDocxPreview(page) {
    docxPreview.innerHTML = "";
    (page.layout && page.layout.lines ? page.layout.lines : [""]).forEach(function (line) {
      const paragraph = document.createElement("p");
      paragraph.className = "docx-preview-line";
      paragraph.textContent = line || " ";
      docxPreview.appendChild(paragraph);
    });
  }

  async function renderSelectedPageBackground() {
    const page = getSelectedPage();
    hideAllBackgrounds();

    if (!page) {
      return;
    }

    if (page.kind === "pdf") {
      pdfCanvas.hidden = false;
      const token = ++state.previewToken;
      try {
        await renderPdfPageToCanvas(page);
        if (token !== state.previewToken) {
          return;
        }
      } catch (error) {
        console.error(error);
        hideAllBackgrounds();
        blankPreview.hidden = false;
        blankPreview.querySelector("span").textContent =
          "PDF preview could not be shown, but export is still available.";
      }
      return;
    }

    if (page.kind === "image") {
      const record = getDocumentById(page.documentId);
      if (record) {
        imagePreview.hidden = false;
        imagePreview.src = record.objectUrl;
      }
      return;
    }

    if (page.kind === "docx") {
      docxPreview.hidden = false;
      renderDocxPreview(page);
      return;
    }

    blankPreview.hidden = false;
    blankPreview.querySelector("span").textContent = "Blank page ready for fields";
  }

  function renderEditor() {
    const page = getSelectedPage();
    const hasPage = Boolean(page);

    editorEmpty.hidden = hasPage;
    editorStageWrap.hidden = !hasPage;

    if (!page) {
      hideAllBackgrounds();
      return;
    }

    pageFrame.style.setProperty("--page-ratio", String(page.width / page.height));
    renderSelectedPageBackground();
    renderFieldLayer();
  }

  function render() {
    ensureSelection();
    renderPageList();
    renderEditor();
    renderInspector();
    renderSignaturePreview();
    setToolAvailability();
  }

  function createField(kind) {
    const page = getSelectedPage();
    if (!page) {
      return null;
    }

    const sameKindCount = page.fields.filter(function (field) {
      return field.kind === kind;
    }).length;

    const defaults = {
      text: {
        label: "Short answer " + (sameKindCount + 1),
        width: 0.48,
        height: 0.085,
      },
      textarea: {
        label: "Paragraph box " + (sameKindCount + 1),
        width: 0.62,
        height: 0.2,
      },
      signature: {
        label: "Signature " + (sameKindCount + 1),
        width: 0.28,
        height: 0.12,
      },
      checkbox: {
        label: "Checkbox " + (sameKindCount + 1),
        width: 0.06,
        height: 0.06,
      },
    };

    const config = defaults[kind];
    const nextY = clamp(0.14 + page.fields.length * 0.08, 0.12, 0.82);

    return {
      id: state.nextFieldId++,
      kind: kind,
      label: config.label,
      x: 0.12,
      y: nextY,
      width: config.width,
      height: config.height,
    };
  }

  function addField(kind) {
    const page = getSelectedPage();
    if (!page) {
      setStatus("Choose a page first, then add a field.");
      return;
    }
    const field = createField(kind);
    page.fields.push(field);
    state.selectedFieldId = field.id;
    resetDownloadLink();
    render();
    setStatus(formatFieldType(kind) + " added. Drag it into place.");
  }

  function removeSelectedField() {
    const page = getSelectedPage();
    const field = getSelectedField();
    if (!page || !field) {
      return;
    }
    page.fields = page.fields.filter(function (entry) {
      return entry.id !== field.id;
    });
    state.selectedFieldId = null;
    resetDownloadLink();
    render();
    setStatus("Selected field removed.");
  }

  function addBlankPage() {
    const page = createBlankPage();
    const selectedIndex = state.pages.findIndex(function (entry) {
      return entry.id === state.selectedPageId;
    });

    if (selectedIndex >= 0) {
      state.pages.splice(selectedIndex + 1, 0, page);
    } else {
      state.pages.push(page);
    }

    state.selectedPageId = page.id;
    state.selectedFieldId = null;
    resetDownloadLink();
    render();
    setStatus("Blank page added.");
  }

  function removePage(pageId) {
    const index = state.pages.findIndex(function (entry) {
      return entry.id === pageId;
    });
    if (index < 0) {
      return;
    }
    state.pages.splice(index, 1);
    state.selectedPageId = state.pages[index]
      ? state.pages[index].id
      : state.pages[index - 1]
      ? state.pages[index - 1].id
      : null;
    state.selectedFieldId = null;
    resetDownloadLink();
    render();
    setStatus("Page removed.");
  }

  function movePage(pageId, direction) {
    const index = state.pages.findIndex(function (entry) {
      return entry.id === pageId;
    });
    const targetIndex = direction === "up" ? index - 1 : index + 1;
    if (index < 0 || targetIndex < 0 || targetIndex >= state.pages.length) {
      return;
    }
    const page = state.pages[index];
    state.pages.splice(index, 1);
    state.pages.splice(targetIndex, 0, page);
    resetDownloadLink();
    render();
    setStatus("Page order saved.");
  }

  function startInteraction(event, mode, fieldId) {
    const page = getSelectedPage();
    if (!page) return;
    const field = page.fields.find(function (entry) {
      return entry.id === fieldId;
    });
    if (!field) return;

    const rect = fieldLayer.getBoundingClientRect();
    if (!rect.width || !rect.height) return;

    state.interaction = {
      pointerId: event.pointerId,
      mode: mode,
      fieldId: fieldId,
      startX: event.clientX,
      startY: event.clientY,
      initialX: field.x,
      initialY: field.y,
      initialWidth: field.width,
      initialHeight: field.height,
      layerRect: rect,
    };

    fieldLayer.setPointerCapture(event.pointerId);
    event.preventDefault();
  }

  function applyInteraction(event) {
    if (!state.interaction || state.interaction.pointerId !== event.pointerId) {
      return;
    }

    const page = getSelectedPage();
    if (!page) {
      return;
    }

    const field = page.fields.find(function (entry) {
      return entry.id === state.interaction.fieldId;
    });
    if (!field) {
      return;
    }

    const dx = (event.clientX - state.interaction.startX) / state.interaction.layerRect.width;
    const dy = (event.clientY - state.interaction.startY) / state.interaction.layerRect.height;

    if (state.interaction.mode === "move") {
      field.x = clamp(state.interaction.initialX + dx, 0.02, 0.98 - field.width);
      field.y = clamp(state.interaction.initialY + dy, 0.03, 0.97 - field.height);
    } else {
      const minSize = field.kind === "checkbox" ? 0.05 : 0.08;
      if (field.kind === "checkbox") {
        const nextSize = clamp(
          Math.max(state.interaction.initialWidth + dx, state.interaction.initialHeight + dy),
          minSize,
          Math.min(0.18, 0.98 - field.x, 0.97 - field.y)
        );
        field.width = nextSize;
        field.height = nextSize;
      } else {
        field.width = clamp(state.interaction.initialWidth + dx, minSize, 0.98 - field.x);
        field.height = clamp(
          state.interaction.initialHeight + dy,
          field.kind === "textarea" ? 0.14 : field.kind === "signature" ? 0.09 : 0.07,
          0.97 - field.y
        );
      }
    }

    renderFieldLayer();
    renderInspector();
    resetDownloadLink();
  }

  function endInteraction(event) {
    if (!state.interaction || state.interaction.pointerId !== event.pointerId) {
      return;
    }
    if (fieldLayer.hasPointerCapture(event.pointerId)) {
      fieldLayer.releasePointerCapture(event.pointerId);
    }
    state.interaction = null;
    setStatus("Field position saved.");
  }

  function toPdfRect(page, field) {
    const x = field.x * page.width;
    const width = field.width * page.width;
    const height = field.height * page.height;
    const y = page.height - (field.y + field.height) * page.height;
    return { x: x, y: y, width: width, height: height };
  }

  function drawDocxPage(pdfPage, page, baseFont) {
    const layout = page.layout;
    if (!layout) {
      return;
    }

    const left = DOCX_LAYOUT.marginX;
    let y = page.height - DOCX_LAYOUT.marginY - DOCX_LAYOUT.fontSize;

    pdfPage.drawRectangle({
      x: 0,
      y: 0,
      width: page.width,
      height: page.height,
      color: rgb(1, 1, 1),
    });

    layout.lines.forEach(function (line) {
      if (line) {
        pdfPage.drawText(line, {
          x: left,
          y: y,
          size: DOCX_LAYOUT.fontSize,
          font: baseFont,
          color: rgb(0.16, 0.19, 0.24),
        });
      }
      y -= DOCX_LAYOUT.lineHeight;
    });
  }

  function drawFieldLabels(pdfPage, field, rect, boldFont) {
    const label = field.label || formatFieldType(field.kind);
    if (field.kind === "checkbox") {
      const labelY = rect.y + rect.height * 0.2;
      pdfPage.drawText(label, {
        x: rect.x + rect.width + 8,
        y: labelY,
        size: 10,
        font: boldFont,
        color: rgb(0.13, 0.19, 0.26),
      });
      return;
    }

    const labelY = Math.min(pdfPage.getHeight() - 12, rect.y + rect.height + 6);
    pdfPage.drawText(label, {
      x: rect.x,
      y: labelY,
      size: 9,
      font: boldFont,
      color: rgb(0.13, 0.19, 0.26),
    });
  }

  function fitInsideRect(width, height, targetWidth, targetHeight) {
    const scale = Math.min(targetWidth / width, targetHeight / height);
    return {
      width: width * scale,
      height: height * scale,
    };
  }

  function drawSignatureField(outputPage, field, rect, boldFont, signatureImage) {
    drawFieldLabels(outputPage, field, rect, boldFont);

    const lineY = rect.y + rect.height * 0.2;
    outputPage.drawLine({
      start: { x: rect.x, y: lineY },
      end: { x: rect.x + rect.width, y: lineY },
      thickness: 1,
      color: rgb(0.32, 0.4, 0.5),
    });

    if (signatureImage) {
      const availableWidth = Math.max(10, rect.width * 0.94);
      const availableHeight = Math.max(10, rect.height * 0.62);
      const size = fitInsideRect(signatureImage.width, signatureImage.height, availableWidth, availableHeight);
      outputPage.drawImage(signatureImage.image, {
        x: rect.x + (rect.width - size.width) / 2,
        y: rect.y + rect.height * 0.3,
        width: size.width,
        height: size.height,
      });
    } else if (state.signatureName) {
      outputPage.drawText(state.signatureName, {
        x: rect.x + 4,
        y: rect.y + rect.height * 0.46,
        size: Math.min(18, Math.max(10, rect.height * 0.34)),
        font: boldFont,
        color: rgb(0.17, 0.24, 0.34),
      });
    }
  }

  async function buildFillablePdf() {
    if (!state.pages.length) {
      return;
    }

    buildBtn.disabled = true;
    setStatus("Building your fillable PDF...");

    try {
      const outputPdf = await PDFDocument.create();
      const form = outputPdf.getForm();
      const baseFont = await outputPdf.embedFont(StandardFonts.Helvetica);
      const boldFont = await outputPdf.embedFont(StandardFonts.HelveticaBold);
      const sourcePdfCache = new Map();
      const imageCache = new Map();
      const signatureImageCache = new Map();

      for (const page of state.pages) {
        let outputPage;

        if (page.kind === "pdf") {
          let sourcePdf = sourcePdfCache.get(page.documentId);
          if (!sourcePdf) {
            const sourceRecord = getDocumentById(page.documentId);
            sourcePdf = await PDFDocument.load(sourceRecord.bytes);
            sourcePdfCache.set(page.documentId, sourcePdf);
          }
          const copiedPages = await outputPdf.copyPages(sourcePdf, [page.sourcePageIndex]);
          outputPage = outputPdf.addPage(copiedPages[0]);
        } else if (page.kind === "image") {
          outputPage = outputPdf.addPage([page.width, page.height]);
          outputPage.drawRectangle({
            x: 0,
            y: 0,
            width: page.width,
            height: page.height,
            color: rgb(1, 1, 1),
          });

          const imageRecord = getDocumentById(page.documentId);
          let embeddedImage = imageCache.get(page.documentId);
          if (!embeddedImage) {
            embeddedImage =
              imageRecord.mimeType === "image/png"
                ? await outputPdf.embedPng(imageRecord.bytes)
                : await outputPdf.embedJpg(imageRecord.bytes);
            imageCache.set(page.documentId, embeddedImage);
          }

          const placement = fitIntoPage(
            imageRecord.imageWidth,
            imageRecord.imageHeight,
            page.width,
            page.height,
            24
          );
          outputPage.drawImage(embeddedImage, placement);
        } else if (page.kind === "docx") {
          outputPage = outputPdf.addPage([page.width, page.height]);
          drawDocxPage(outputPage, page, baseFont);
        } else {
          outputPage = outputPdf.addPage([page.width, page.height]);
          outputPage.drawRectangle({
            x: 0,
            y: 0,
            width: page.width,
            height: page.height,
            color: rgb(1, 1, 1),
          });
        }

        for (const field of page.fields) {
          const rect = toPdfRect(page, field);
          const fieldName =
            "page-" + page.id + "-" + slugify(field.label || field.kind) + "-" + String(field.id);

          if (field.kind === "checkbox") {
            drawFieldLabels(outputPage, field, rect, boldFont);
            const size = Math.min(rect.width, rect.height);
            const checkbox = form.createCheckBox(fieldName);
            checkbox.addToPage(outputPage, {
              x: rect.x,
              y: rect.y + (rect.height - size) / 2,
              width: size,
              height: size,
              borderWidth: 1,
              borderColor: rgb(0.19, 0.27, 0.37),
              backgroundColor: rgb(1, 1, 1),
            });
            return;
          }

          if (field.kind === "signature") {
            let signatureImage = null;
            if (state.signatureBytes && state.signatureMimeType) {
              const cacheKey = state.signatureMimeType;
              signatureImage = signatureImageCache.get(cacheKey) || null;
              if (!signatureImage) {
                const image =
                  state.signatureMimeType === "image/png"
                    ? await outputPdf.embedPng(state.signatureBytes)
                    : await outputPdf.embedJpg(state.signatureBytes);
                signatureImage = {
                  image: image,
                  width: image.width,
                  height: image.height,
                };
                signatureImageCache.set(cacheKey, signatureImage);
              }
            }
            drawSignatureField(outputPage, field, rect, boldFont, signatureImage);
            return;
          }

          const textField = form.createTextField(fieldName);
          if (field.kind === "textarea") {
            textField.enableMultiline();
          }
          drawFieldLabels(outputPage, field, rect, boldFont);
          textField.addToPage(outputPage, {
            x: rect.x,
            y: rect.y,
            width: rect.width,
            height: rect.height,
            font: baseFont,
            borderWidth: 1,
            borderColor: rgb(0.19, 0.27, 0.37),
            backgroundColor: rgb(1, 1, 1),
            textColor: rgb(0.1, 0.14, 0.19),
          });
        }
      }

      const bytes = await outputPdf.save();
      const blob = new Blob([bytes], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const dateStamp = new Date().toISOString().slice(0, 10);
      setDownloadLink(url, "swiftpdf-fillable-form-" + dateStamp + ".pdf");
      setStatus("Done. Your fillable PDF is ready to download.");
    } catch (error) {
      console.error(error);
      setStatus("We could not build that file. Please try again with a simpler document.");
    } finally {
      buildBtn.disabled = state.pages.length === 0;
    }
  }

  browseBtn.addEventListener("click", function () {
    fileInput.click();
  });

  fileInput.addEventListener("change", function (event) {
    addFiles(event.target.files);
    fileInput.value = "";
  });

  let dragDepth = 0;

  dropzone.addEventListener("dragenter", function (event) {
    event.preventDefault();
    dragDepth += 1;
    dropzone.classList.add("active");
  });

  dropzone.addEventListener("dragover", function (event) {
    event.preventDefault();
  });

  dropzone.addEventListener("dragleave", function (event) {
    event.preventDefault();
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) {
      dropzone.classList.remove("active");
    }
  });

  dropzone.addEventListener("drop", function (event) {
    event.preventDefault();
    dragDepth = 0;
    dropzone.classList.remove("active");
    if (event.dataTransfer && event.dataTransfer.files) {
      addFiles(event.dataTransfer.files);
    }
  });

  addBlankPageBtn.addEventListener("click", addBlankPage);
  removePageBtn.addEventListener("click", function () {
    const page = getSelectedPage();
    if (page) {
      removePage(page.id);
    }
  });

  addTextFieldBtn.addEventListener("click", function () {
    addField("text");
  });

  addTextareaFieldBtn.addEventListener("click", function () {
    addField("textarea");
  });

  addCheckboxFieldBtn.addEventListener("click", function () {
    addField("checkbox");
  });

  addSignatureFieldBtn.addEventListener("click", function () {
    addField("signature");
  });

  pageList.addEventListener("click", function (event) {
    const target = event.target.closest("button[data-action]");
    if (!target) {
      return;
    }

    const pageId = Number(target.getAttribute("data-page-id"));
    const action = target.getAttribute("data-action");

    if (action === "select-page") {
      selectPage(pageId);
      return;
    }

    if (action === "move-up") {
      movePage(pageId, "up");
      return;
    }

    if (action === "move-down") {
      movePage(pageId, "down");
      return;
    }

    if (action === "remove-page") {
      removePage(pageId);
    }
  });

  fieldLayer.addEventListener("pointerdown", function (event) {
    const resizeHandle = event.target.closest("[data-action='resize-field']");
    const fieldNode = event.target.closest(".field-box");

    if (!fieldNode) {
      selectField(null);
      return;
    }

    const fieldId = Number(fieldNode.getAttribute("data-field-id"));
    selectField(fieldId);
    startInteraction(event, resizeHandle ? "resize" : "move", fieldId);
  });

  fieldLayer.addEventListener("pointermove", applyInteraction);
  fieldLayer.addEventListener("pointerup", endInteraction);
  fieldLayer.addEventListener("pointercancel", endInteraction);

  fieldLabelInput.addEventListener("input", function () {
    const field = getSelectedField();
    if (!field) {
      return;
    }
    field.label = fieldLabelInput.value.trim() || formatFieldType(field.kind);
    resetDownloadLink();
    renderFieldLayer();
    renderInspector();
  });

  signatureNameInput.addEventListener("input", function () {
    state.signatureName = signatureNameInput.value.trim();
    resetDownloadLink();
    renderSignaturePreview();
    renderFieldLayer();
  });

  signatureUploadInput.addEventListener("change", function () {
    const file = signatureUploadInput.files && signatureUploadInput.files[0];
    if (!file) {
      return;
    }
    if (!file.type || (file.type !== "image/png" && file.type !== "image/jpeg")) {
      setStatus("Please choose a PNG or JPG signature image.");
      signatureUploadInput.value = "";
      return;
    }

    const reader = new FileReader();
    reader.onload = function () {
      const dataUrl = typeof reader.result === "string" ? reader.result : null;
      state.signatureDataUrl = dataUrl;
      state.signatureMimeType = file.type;
      state.signatureBytes = dataUrlToBytes(dataUrl);
      resetDownloadLink();
      renderSignaturePreview();
      renderFieldLayer();
      setStatus("Signature image loaded.");
    };
    reader.onerror = function () {
      setStatus("We could not read that signature image. Please try another one.");
    };
    reader.readAsDataURL(file);
  });

  clearSignatureBtn.addEventListener("click", function () {
    state.signatureDataUrl = null;
    state.signatureMimeType = "";
    state.signatureBytes = null;
    signatureUploadInput.value = "";
    resetDownloadLink();
    renderSignaturePreview();
    renderFieldLayer();
    setStatus("Signature cleared.");
  });

  removeFieldBtn.addEventListener("click", removeSelectedField);
  buildBtn.addEventListener("click", buildFillablePdf);

  downloadLink.addEventListener("click", function (event) {
    if (downloadLink.classList.contains("disabled")) {
      event.preventDefault();
    }
  });

  window.addEventListener("beforeunload", function () {
    resetDownloadLink();
    state.documents.forEach(function (record) {
      if (record.objectUrl) {
        URL.revokeObjectURL(record.objectUrl);
      }
    });
  });

  window.addEventListener("resize", function () {
    const page = getSelectedPage();
    if (page && page.kind === "pdf") {
      renderEditor();
    }
  });

  render();
})();

