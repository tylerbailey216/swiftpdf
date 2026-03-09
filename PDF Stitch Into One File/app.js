const queue = [];
let nextId = 1;
let currentBlobUrl = null;

const fileInput = document.getElementById("fileInput");
const dropzone = document.getElementById("dropzone");
const browseBtn = document.getElementById("browseBtn");
const fileList = document.getElementById("fileList");
const emptyState = document.getElementById("emptyState");
const mergeBtn = document.getElementById("mergeBtn");
const downloadLink = document.getElementById("downloadLink");
const statusEl = document.getElementById("status");
const clearBtn = document.getElementById("clearBtn");
const imagePageSizeSelect = document.getElementById("imagePageSize");
const pdfPageFitSelect = document.getElementById("pdfPageFit");
const watermarkInput = document.getElementById("watermarkText");
const watermarkOpacityInput = document.getElementById("watermarkOpacity");
const watermarkOpacityValue = document.getElementById("watermarkOpacityValue");
const watermarkToneInput = document.getElementById("watermarkTone");
const watermarkToneValue = document.getElementById("watermarkToneValue");
const watermarkSizeInput = document.getElementById("watermarkSize");
const watermarkSizeValue = document.getElementById("watermarkSizeValue");
const watermarkOrientationSelect = document.getElementById("watermarkOrientation");
const footerInput = document.getElementById("footerText");
const footerOpacityInput = document.getElementById("footerOpacity");
const footerOpacityValue = document.getElementById("footerOpacityValue");
const footerToneInput = document.getElementById("footerTone");
const footerToneValue = document.getElementById("footerToneValue");
const signNameInput = document.getElementById("signName");
const signDateInput = document.getElementById("signDate");
const signNoteInput = document.getElementById("signNote");
const signPlacementSelect = document.getElementById("signPlacement");
const signOpacityInput = document.getElementById("signOpacity");
const signOpacityValue = document.getElementById("signOpacityValue");
const signaturePad = document.getElementById("signaturePad");
const signatureClearBtn = document.getElementById("signatureClear");
const signatureUploadInput = document.getElementById("signatureUpload");
const previewOrientationSelect = document.getElementById("previewOrientation");
const previewGrid = document.getElementById("previewGrid");
const previewWatermarks = document.querySelectorAll("[data-preview-watermark]");
const previewFooters = document.querySelectorAll("[data-preview-footer]");
const previewFrames = document.querySelectorAll("[data-preview-frame]");
const previewHandles = document.querySelectorAll("[data-preview-handle]");
const livePreviewSourceSelect = document.getElementById("livePreviewSource");
const livePreviewGrid = document.getElementById("livePreviewGrid");
const livePreviewEmpty = document.getElementById("livePreviewEmpty");
const livePreviewMeta = document.getElementById("livePreviewMeta");
const livePreviewBeforeFrame = document.getElementById("livePreviewBefore");
const livePreviewAfterFrame = document.getElementById("livePreviewAfter");
const watermarkPageList = document.getElementById("watermarkPageList");
const watermarkPageEmpty = document.getElementById("watermarkPageEmpty");
const watermarkSelectAllBtn = document.getElementById("watermarkSelectAll");
const watermarkSelectNoneBtn = document.getElementById("watermarkSelectNone");
const footerPageList = document.getElementById("footerPageList");
const footerPageEmpty = document.getElementById("footerPageEmpty");
const footerSelectAllBtn = document.getElementById("footerSelectAll");
const footerSelectNoneBtn = document.getElementById("footerSelectNone");
const signPageList = document.getElementById("signPageList");
const signPageEmpty = document.getElementById("signPageEmpty");
const signSelectAllBtn = document.getElementById("signSelectAll");
const signSelectNoneBtn = document.getElementById("signSelectNone");

const { PDFDocument, StandardFonts, rgb, degrees } = PDFLib;

const PAGE_SIZES = {
  letter: { width: 612, height: 792 },
  a4: { width: 595.28, height: 841.89 },
};

const watermarkBox = {
  centerX: 0.5,
  centerY: 0.5,
  width: 0.7,
  height: 0.35,
};

const WATERMARK_BOX_MIN_WIDTH = 0.2;
const selectedFooterPages = new Set();
const selectedWatermarkPages = new Set();
const selectedSignPages = new Set();
const knownFooterPageKeys = new Set();
const knownWatermarkPageKeys = new Set();
const knownSignPageKeys = new Set();
let footerPageKeys = [];
let watermarkPageKeys = [];
let signPageKeys = [];
let footerDefaultSelected = true;
let watermarkDefaultSelected = true;
let signDefaultSelected = true;
let livePreviewBeforeUrl = null;
let livePreviewAfterUrl = null;
let livePreviewDebounceTimer = null;
let livePreviewToken = 0;
let signatureDataUrl = null;
let signaturePadContext = null;
let signaturePadDrawing = null;

const previewMeasureContext = (() => {
  if (typeof document === "undefined") return null;
  const canvas = document.createElement("canvas");
  return canvas.getContext("2d");
})();

const clamp = (value, min, max) => Math.min(Math.max(value, min), max);

const formatSize = (bytes) => {
  if (bytes === 0) return "0 KB";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / Math.pow(1024, index);
  return `${value.toFixed(value >= 10 || index === 0 ? 0 : 1)} ${units[index]}`;
};

const formatDateForSignLine = (value) => {
  if (!value) return "";
  const parts = String(value).split("-");
  if (parts.length !== 3) return value;
  return `${parts[1]}/${parts[2]}/${parts[0]}`;
};

const dataUrlToBytes = (dataUrl) => {
  if (!dataUrl || typeof dataUrl !== "string" || !dataUrl.includes(",")) {
    return null;
  }
  const payload = dataUrl.split(",")[1];
  if (!payload) return null;
  const binary = atob(payload);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
};

const getFileKind = (file) => {
  const name = file.name.toLowerCase();
  if (file.type === "application/pdf" || name.endsWith(".pdf")) {
    return "pdf";
  }
  if (
    file.type === "image/jpeg" ||
    name.endsWith(".jpg") ||
    name.endsWith(".jpeg") ||
    name.endsWith(".jfif")
  ) {
    return "jpeg";
  }
  if (file.type === "image/png" || name.endsWith(".png")) {
    return "png";
  }
  return null;
};

const syncSignaturePadCanvasSize = () => {
  if (!signaturePad) return;
  const ratio = window.devicePixelRatio || 1;
  const cssWidth = signaturePad.clientWidth || 560;
  const cssHeight = signaturePad.clientHeight || 150;
  const nextWidth = Math.max(1, Math.round(cssWidth * ratio));
  const nextHeight = Math.max(1, Math.round(cssHeight * ratio));
  if (signaturePad.width === nextWidth && signaturePad.height === nextHeight) {
    return;
  }
  const previous = signatureDataUrl;
  signaturePad.width = nextWidth;
  signaturePad.height = nextHeight;
  if (!signaturePadContext) {
    signaturePadContext = signaturePad.getContext("2d");
  }
  if (!signaturePadContext) return;
  signaturePadContext.setTransform(ratio, 0, 0, ratio, 0, 0);
  signaturePadContext.lineCap = "round";
  signaturePadContext.lineJoin = "round";
  signaturePadContext.lineWidth = 2.2;
  signaturePadContext.strokeStyle = "#d9ebff";
  if (previous) {
    const img = new Image();
    img.onload = () => {
      if (!signaturePadContext) return;
      signaturePadContext.clearRect(0, 0, cssWidth, cssHeight);
      signaturePadContext.drawImage(img, 0, 0, cssWidth, cssHeight);
    };
    img.src = previous;
  } else {
    signaturePadContext.clearRect(0, 0, cssWidth, cssHeight);
  }
};

const saveSignatureFromPad = () => {
  if (!signaturePad) return;
  signatureDataUrl = signaturePad.toDataURL("image/png");
  if (signatureUploadInput) {
    signatureUploadInput.value = "";
  }
};

const clearSignaturePad = () => {
  if (!signaturePad) return;
  const width = signaturePad.clientWidth || 560;
  const height = signaturePad.clientHeight || 150;
  if (!signaturePadContext) {
    signaturePadContext = signaturePad.getContext("2d");
  }
  if (signaturePadContext) {
    signaturePadContext.clearRect(0, 0, width, height);
  }
  signatureDataUrl = null;
};

const loadSignatureDataUrlToPad = (dataUrl) => {
  if (!signaturePad || !dataUrl) return;
  const width = signaturePad.clientWidth || 560;
  const height = signaturePad.clientHeight || 150;
  if (!signaturePadContext) {
    signaturePadContext = signaturePad.getContext("2d");
  }
  if (!signaturePadContext) return;
  const image = new Image();
  image.onload = () => {
    signaturePadContext.clearRect(0, 0, width, height);
    const fit = Math.min(width / image.width, height / image.height);
    const drawWidth = image.width * fit;
    const drawHeight = image.height * fit;
    const drawX = (width - drawWidth) / 2;
    const drawY = (height - drawHeight) / 2;
    signaturePadContext.drawImage(image, drawX, drawY, drawWidth, drawHeight);
  };
  image.src = dataUrl;
};

const getSignaturePointerPoint = (event) => {
  if (!signaturePad) return null;
  const rect = signaturePad.getBoundingClientRect();
  return {
    x: event.clientX - rect.left,
    y: event.clientY - rect.top,
  };
};

const startSignatureDraw = (event) => {
  if (!signaturePad || !signaturePadContext) return;
  const point = getSignaturePointerPoint(event);
  if (!point) return;
  signaturePadDrawing = {
    pointerId: event.pointerId,
    x: point.x,
    y: point.y,
  };
  signaturePad.setPointerCapture(event.pointerId);
  event.preventDefault();
};

const moveSignatureDraw = (event) => {
  if (!signaturePad || !signaturePadContext || !signaturePadDrawing) return;
  if (event.pointerId !== signaturePadDrawing.pointerId) return;
  const point = getSignaturePointerPoint(event);
  if (!point) return;
  signaturePadContext.beginPath();
  signaturePadContext.moveTo(signaturePadDrawing.x, signaturePadDrawing.y);
  signaturePadContext.lineTo(point.x, point.y);
  signaturePadContext.stroke();
  signaturePadDrawing.x = point.x;
  signaturePadDrawing.y = point.y;
};

const endSignatureDraw = (event) => {
  if (!signaturePad || !signaturePadDrawing) return;
  if (event.pointerId !== signaturePadDrawing.pointerId) return;
  if (signaturePad.hasPointerCapture(event.pointerId)) {
    signaturePad.releasePointerCapture(event.pointerId);
  }
  signaturePadDrawing = null;
  saveSignatureFromPad();
  resetDownloadLink();
  scheduleWatermarkPreviewUpdate();
  updateStatus("Signature saved.");
};

const getSignSettings = () => {
  const name = signNameInput?.value.trim() ?? "";
  const rawDate = signDateInput?.value || "";
  const note = signNoteInput?.value.trim() ?? "";
  const placement = signPlacementSelect?.value || "bottom-right";
  const opacity = signOpacityInput ? Number(signOpacityInput.value) : 0.95;
  const signDate = rawDate ? formatDateForSignLine(rawDate) : "";
  const hasContent = Boolean(name || signDate || note || signatureDataUrl);
  return {
    name,
    signDate,
    note,
    placement,
    opacity,
    hasContent,
    signatureDataUrl,
  };
};

const resolveImagePageSize = async () => {
  const preference = imagePageSizeSelect?.value ?? "match";
  if (preference === "letter") {
    return { ...PAGE_SIZES.letter };
  }
  if (preference === "a4") {
    return { ...PAGE_SIZES.a4 };
  }

  const pdfItem = queue.find((item) => item.kind === "pdf");
  if (!pdfItem) {
    return { ...PAGE_SIZES.letter };
  }
  const cachedSize = getPdfItemSize(pdfItem);
  if (cachedSize) {
    return { ...cachedSize };
  }

  try {
    const bytes = await pdfItem.file.arrayBuffer();
    const pdfDoc = await PDFDocument.load(bytes);
    const firstPage = pdfDoc.getPage(0);
    const { mediaBox, trimmedBox } = getPdfPageBoxes(firstPage);
    const targetBox = getPdfPageFitMode() === "trim" ? trimmedBox : mediaBox;
    if (targetBox) {
      return { width: targetBox.width, height: targetBox.height };
    }
    const { width, height } = firstPage.getSize();
    return { width, height };
  } catch (error) {
    console.error(error);
    return { ...PAGE_SIZES.letter };
  }
};

const fitImageToPage = (imageWidth, imageHeight, pageWidth, pageHeight) => {
  const scale = Math.min(pageWidth / imageWidth, pageHeight / imageHeight);
  const width = imageWidth * scale;
  const height = imageHeight * scale;
  return {
    x: (pageWidth - width) / 2,
    y: (pageHeight - height) / 2,
    width,
    height,
  };
};

const PAGE_BOX_EPSILON = 0.5;

const normalizePageBox = (box) => {
  if (!box) return null;
  const x = Number(box.x ?? 0);
  const y = Number(box.y ?? 0);
  const width = Number(box.width);
  const height = Number(box.height);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return null;
  }
  return { x, y, width, height };
};

const sizeToBox = (size) => {
  if (!size) return null;
  const width = Number(size.width);
  const height = Number(size.height);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return null;
  }
  return { x: 0, y: 0, width, height };
};

const isBoxSmallerThan = (candidate, reference) => {
  if (!candidate || !reference) return false;
  return (
    candidate.width < reference.width - PAGE_BOX_EPSILON ||
    candidate.height < reference.height - PAGE_BOX_EPSILON
  );
};

const getPdfPageFitMode = () => pdfPageFitSelect?.value ?? "original";

const getPdfPageBoxes = (page) => {
  const mediaBox = normalizePageBox(page.getMediaBox?.());
  const cropBox = normalizePageBox(page.getCropBox?.());
  const trimBox = normalizePageBox(page.getTrimBox?.());
  let trimmedBox = null;

  [cropBox, trimBox].forEach((box) => {
    if (!box || !mediaBox) return;
    if (!isBoxSmallerThan(box, mediaBox)) return;
    if (!trimmedBox || box.width * box.height < trimmedBox.width * trimmedBox.height) {
      trimmedBox = box;
    }
  });

  return { mediaBox, trimmedBox: trimmedBox || mediaBox };
};

const getPdfItemSize = (item) => {
  if (!item) return null;
  const mode = getPdfPageFitMode();
  if (mode === "trim" && item.pageTrimSize) {
    return item.pageTrimSize;
  }
  return item.pageSize;
};

const toBoundingBox = (box) => {
  if (!box) return null;
  return {
    left: box.x,
    bottom: box.y,
    right: box.x + box.width,
    top: box.y + box.height,
  };
};

const getWatermarkText = () => {
  const text = watermarkInput?.value.trim();
  return text ? text : null;
};

const getFooterText = () => {
  const text = footerInput?.value.trim();
  return text ? text : null;
};

const getWatermarkAngle = () => {
  const orientation = watermarkOrientationSelect?.value ?? "diagonal";
  if (orientation === "straight") {
    return 0;
  }
  if (orientation === "inverse") {
    return 30;
  }
  return -30;
};

const getWatermarkSettings = () => {
  const opacity = watermarkOpacityInput ? Number(watermarkOpacityInput.value) : 0.22;
  const tonePercent = watermarkToneInput ? Number(watermarkToneInput.value) : 60;
  const tone = 0.1 + 0.8 * (tonePercent / 100);
  const toneRgb = Math.round(tone * 255);
  const sizePercent = watermarkSizeInput ? Number(watermarkSizeInput.value) : 100;
  const sizeScale = Number.isFinite(sizePercent) ? sizePercent / 100 : 1;
  const angle = getWatermarkAngle();
  return { opacity, tone, tonePercent, toneRgb, sizePercent, sizeScale, angle };
};

const getFooterSettings = () => {
  const opacity = footerOpacityInput ? Number(footerOpacityInput.value) : 0.7;
  const tonePercent = footerToneInput ? Number(footerToneInput.value) : 13;
  const tone = 0.1 + 0.8 * (tonePercent / 100);
  const toneRgb = Math.round(tone * 255);
  return { opacity, tone, tonePercent, toneRgb };
};

const getPreviewWatermarkText = () => {
  const text = watermarkInput?.value.trim();
  if (text) return text;
  return watermarkInput?.placeholder?.trim() || "Sample watermark";
};

const getPreviewFooterText = () => {
  const text = footerInput?.value.trim();
  if (text) return text;
  return footerInput?.placeholder?.trim() || "Sample bottom note";
};

const getPreviewSignName = () => {
  const text = signNameInput?.value.trim();
  if (text) return text;
  return signNameInput?.placeholder?.trim() || "Signer Name";
};

const getReferencePageSize = () => {
  const pdfItem = queue.find((item) => item.kind === "pdf" && item.pageSize);
  const pdfSize = getPdfItemSize(pdfItem);
  if (pdfSize) {
    return pdfSize;
  }
  const preference = imagePageSizeSelect?.value ?? "letter";
  if (preference === "a4") {
    return PAGE_SIZES.a4;
  }
  return PAGE_SIZES.letter;
};

const getPreviewAspectRatio = () => {
  const size = getReferencePageSize();
  let ratio = size.width / size.height;
  const orientation = previewOrientationSelect?.value ?? "auto";
  if (orientation === "portrait" && ratio > 1) {
    ratio = 1 / ratio;
  } else if (orientation === "landscape" && ratio < 1) {
    ratio = 1 / ratio;
  }
  return ratio;
};

const updatePreviewAspect = () => {
  if (!previewGrid) return;
  const ratio = getPreviewAspectRatio();
  previewGrid.style.setProperty("--preview-aspect", ratio.toFixed(3));
};

const getPreviewFontSize = (boxWidthPx, boxHeightPx, text) => {
  if (!previewMeasureContext || !boxWidthPx || !boxHeightPx) {
    return Math.max(12, Math.min(boxWidthPx, boxHeightPx) * 0.4);
  }
  const safeText = text || "W";
  previewMeasureContext.font = "700 1px 'Helvetica Neue', Arial, sans-serif";
  const metrics = previewMeasureContext.measureText(safeText);
  const widthPerPx = metrics.width || 1;
  const heightPerPx =
    metrics.actualBoundingBoxAscent !== undefined && metrics.actualBoundingBoxDescent !== undefined
      ? metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent
      : 1.15;
  const sizeByWidth = (boxWidthPx * 0.92) / widthPerPx;
  const sizeByHeight = (boxHeightPx * 0.8) / heightPerPx;
  return Math.max(10, Math.min(sizeByWidth, sizeByHeight));
};

const updateWatermarkPreview = () => {
  if (!previewGrid || previewWatermarks.length === 0) return;
  updatePreviewAspect();
  const { opacity, tonePercent, toneRgb, sizePercent, sizeScale, angle } = getWatermarkSettings();
  const left = (watermarkBox.centerX - watermarkBox.width / 2) * 100;
  const top = (watermarkBox.centerY - watermarkBox.height / 2) * 100;
  const width = watermarkBox.width * 100;
  const height = watermarkBox.height * 100;
  previewGrid.style.setProperty("--wm-color", `${toneRgb}, ${toneRgb}, ${toneRgb}`);
  previewGrid.style.setProperty("--wm-opacity", opacity.toFixed(2));
  previewGrid.style.setProperty("--wm-angle", `${angle}deg`);
  previewGrid.style.setProperty("--wm-left", `${left}%`);
  previewGrid.style.setProperty("--wm-top", `${top}%`);
  previewGrid.style.setProperty("--wm-width", `${width}%`);
  previewGrid.style.setProperty("--wm-height", `${height}%`);
  const previewText = getPreviewWatermarkText();
  const previewCanvas = previewGrid.querySelector(".preview-canvas");
  if (previewCanvas) {
    const rect = previewCanvas.getBoundingClientRect();
    const baseFontSize = getPreviewFontSize(
      rect.width * watermarkBox.width,
      rect.height * watermarkBox.height,
      previewText
    );
    const fontSize = Math.max(6, baseFontSize * sizeScale);
    previewGrid.style.setProperty("--wm-font-size", `${fontSize}px`);
  }
  previewWatermarks.forEach((node) => {
    node.textContent = previewText;
  });
  if (watermarkOpacityValue) {
    watermarkOpacityValue.textContent = `${Math.round(opacity * 100)}%`;
  }
  if (watermarkToneValue) {
    watermarkToneValue.textContent = `${tonePercent}%`;
  }
  if (watermarkSizeValue) {
    watermarkSizeValue.textContent = `${Math.round(sizePercent)}%`;
  }

  if (previewFooters.length) {
    const footerText = getPreviewFooterText();
    const hasFooter = Boolean(footerInput?.value.trim());
    const { opacity: footerOpacity, toneRgb: footerToneRgb, tonePercent: footerTonePercent } =
      getFooterSettings();
    const previewOpacity = footerText
      ? hasFooter
        ? footerOpacity
        : Math.min(footerOpacity, 0.35)
      : 0;

    previewFooters.forEach((node) => {
      node.textContent = footerText;
      node.style.opacity = String(previewOpacity);
      node.style.color = `rgb(${footerToneRgb}, ${footerToneRgb}, ${footerToneRgb})`;
    });
    if (footerOpacityValue) {
      footerOpacityValue.textContent = `${Math.round(footerOpacity * 100)}%`;
    }
    if (footerToneValue) {
      footerToneValue.textContent = `${footerTonePercent}%`;
    }
  }

  if (signOpacityValue) {
    const { opacity } = getSignSettings();
    signOpacityValue.textContent = `${Math.round(opacity * 100)}%`;
  }
};

let previewUpdatePending = false;

const scheduleWatermarkPreviewUpdate = () => {
  if (previewUpdatePending) return;
  previewUpdatePending = true;
  window.requestAnimationFrame(() => {
    previewUpdatePending = false;
    updateWatermarkPreview();
  });
};

let watermarkDragState = null;

const startWatermarkDrag = (event, mode) => {
  const frame = event.currentTarget.closest(".preview-frame");
  const canvas = frame?.closest(".preview-canvas");
  if (!frame || !canvas) return;
  const rect = canvas.getBoundingClientRect();
  watermarkDragState = {
    mode,
    pointerId: event.pointerId,
    frame,
    rect,
    startX: event.clientX,
    startY: event.clientY,
    startCenterX: watermarkBox.centerX,
    startCenterY: watermarkBox.centerY,
    startWidth: watermarkBox.width,
    startHeight: watermarkBox.height,
    didMove: false,
  };
  frame.setPointerCapture(event.pointerId);
  event.preventDefault();
  window.addEventListener("pointermove", handleWatermarkDragMove);
  window.addEventListener("pointerup", handleWatermarkDragEnd);
  window.addEventListener("pointercancel", handleWatermarkDragEnd);
};

const handleWatermarkDragMove = (event) => {
  if (!watermarkDragState || event.pointerId !== watermarkDragState.pointerId) return;
  const { rect, mode } = watermarkDragState;
  const dx = (event.clientX - watermarkDragState.startX) / rect.width;
  const dy = (event.clientY - watermarkDragState.startY) / rect.height;
  if (Math.abs(dx) > 0.001 || Math.abs(dy) > 0.001) {
    watermarkDragState.didMove = true;
  }

  if (mode === "move") {
    const minX = watermarkBox.width / 2;
    const minY = watermarkBox.height / 2;
    watermarkBox.centerX = clamp(watermarkDragState.startCenterX + dx, minX, 1 - minX);
    watermarkBox.centerY = clamp(watermarkDragState.startCenterY + dy, minY, 1 - minY);
  } else if (mode === "resize") {
    const ratio = watermarkDragState.startHeight / watermarkDragState.startWidth;
    const delta = (dx + dy) / 2;
    const maxWidthByX = 2 * Math.min(watermarkDragState.startCenterX, 1 - watermarkDragState.startCenterX);
    const maxHeightByY =
      2 * Math.min(watermarkDragState.startCenterY, 1 - watermarkDragState.startCenterY);
    const maxWidthByY = maxHeightByY / ratio;
    const maxWidth = Math.min(1, maxWidthByX, maxWidthByY);
    const nextWidth = clamp(
      watermarkDragState.startWidth + delta,
      WATERMARK_BOX_MIN_WIDTH,
      maxWidth
    );
    watermarkBox.width = nextWidth;
    watermarkBox.height = nextWidth * ratio;
  }

  updateWatermarkPreview();
};

const handleWatermarkDragEnd = (event) => {
  if (!watermarkDragState || event.pointerId !== watermarkDragState.pointerId) return;
  watermarkDragState.frame.releasePointerCapture(event.pointerId);
  window.removeEventListener("pointermove", handleWatermarkDragMove);
  window.removeEventListener("pointerup", handleWatermarkDragEnd);
  window.removeEventListener("pointercancel", handleWatermarkDragEnd);
  if (watermarkDragState.didMove) {
    resetDownloadLink();
    updateStatus("Watermark position saved.");
  }
  watermarkDragState = null;
};

const applyWatermark = (page, text, font, settings, box) => {
  if (!text || !font) return;
  const { width, height } = page.getSize();
  const targetBox = box || watermarkBox;
  const boxWidth = Math.max(1, width * targetBox.width);
  const boxHeight = Math.max(1, height * targetBox.height);
  const widthAtSize = font.widthOfTextAtSize(text, 1);
  const heightAtSize = font.heightAtSize(1);
  const sizeByWidth = (boxWidth * 0.92) / widthAtSize;
  const sizeByHeight = (boxHeight * 0.8) / heightAtSize;
  const baseFontSize = Math.max(10, Math.min(sizeByWidth, sizeByHeight));
  const sizeScale = settings?.sizeScale ?? 1;
  const fontSize = Math.max(6, baseFontSize * sizeScale);
  const textWidth = font.widthOfTextAtSize(text, fontSize);
  const textHeight = font.heightAtSize(fontSize);
  const centerX = width * targetBox.centerX;
  const centerY = height * (1 - targetBox.centerY);
  const angle = settings?.angle ?? -30;
  const radians = (Math.PI / 180) * angle;
  const x = centerX - (textWidth / 2) * Math.cos(radians) + (textHeight / 2) * Math.sin(radians);
  const y = centerY - (textWidth / 2) * Math.sin(radians) - (textHeight / 2) * Math.cos(radians);
  const tone = settings?.tone ?? 0.12;
  const opacity = settings?.opacity ?? 0.22;

  page.drawText(text, {
    x,
    y,
    size: fontSize,
    font,
    color: rgb(tone, tone, tone),
    opacity,
    rotate: degrees(angle),
  });
};

const applyFooter = (page, text, font, settings) => {
  if (!text || !font) return;
  const { width, height } = page.getSize();
  let fontSize = Math.max(9, Math.min(14, Math.min(width, height) * 0.02));
  let textWidth = font.widthOfTextAtSize(text, fontSize);
  const maxTextWidth = Math.max(80, width - 48);
  if (textWidth > maxTextWidth) {
    fontSize *= maxTextWidth / textWidth;
    textWidth = font.widthOfTextAtSize(text, fontSize);
  }
  const x = (width - textWidth) / 2;
  const y = Math.max(12, fontSize * 0.8);
  const tone = settings?.tone ?? 0.2;
  const opacity = settings?.opacity ?? 0.7;

  page.drawText(text, {
    x,
    y,
    size: fontSize,
    font,
    color: rgb(tone, tone, tone),
    opacity,
  });
};

const applyFillAndSign = (page, signInfo, fonts, signatureImage) => {
  if (!signInfo?.hasContent) return;
  const { width, height } = page.getSize();
  const margin = Math.max(18, Math.min(width, height) * 0.03);
  const blockWidth = Math.min(width * 0.44, 250);

  const lines = [];
  if (signInfo.name) lines.push({ text: signInfo.name, kind: "name" });
  if (signInfo.signDate) lines.push({ text: `Date: ${signInfo.signDate}`, kind: "meta" });
  if (signInfo.note) lines.push({ text: signInfo.note, kind: "meta" });

  let signatureHeight = 0;
  let signatureWidth = 0;
  if (signatureImage) {
    const native = signatureImage.scale(1);
    if (native.width > 0 && native.height > 0) {
      signatureWidth = Math.min(blockWidth - 22, 170);
      signatureHeight = signatureWidth * (native.height / native.width);
      if (signatureHeight > 42) {
        const scale = 42 / signatureHeight;
        signatureHeight = 42;
        signatureWidth *= scale;
      }
    }
  }

  const lineHeight = 12;
  const textHeight = lines.length ? lines.length * lineHeight + 4 : 0;
  const blockHeight = Math.max(44, signatureHeight + textHeight + 16);

  let x = width - margin - blockWidth;
  let y = margin;
  if (signInfo.placement === "bottom-left") {
    x = margin;
  } else if (signInfo.placement === "top-right") {
    x = width - margin - blockWidth;
    y = height - margin - blockHeight;
  } else if (signInfo.placement === "center") {
    x = (width - blockWidth) / 2;
    y = (height - blockHeight) / 2;
  }

  page.drawRectangle({
    x,
    y,
    width: blockWidth,
    height: blockHeight,
    color: rgb(1, 1, 1),
    opacity: Math.min(0.24, signInfo.opacity * 0.26),
    borderWidth: 0.7,
    borderColor: rgb(0.72, 0.72, 0.72),
    borderOpacity: Math.min(0.5, signInfo.opacity * 0.45),
  });

  let cursorY = y + blockHeight - 10;
  if (signatureImage && signatureWidth > 0 && signatureHeight > 0) {
    const drawX = x + 10;
    const drawY = cursorY - signatureHeight;
    page.drawImage(signatureImage, {
      x: drawX,
      y: drawY,
      width: signatureWidth,
      height: signatureHeight,
      opacity: signInfo.opacity,
    });
    cursorY = drawY - 6;
  }

  lines.forEach((line, index) => {
    const isName = line.kind === "name";
    const font = isName ? fonts?.nameFont || fonts?.baseFont : fonts?.baseFont;
    if (!font) return;
    const fontSize = isName ? 10.5 : 9;
    page.drawText(line.text, {
      x: x + 10,
      y: cursorY - fontSize,
      size: fontSize,
      font,
      color: rgb(0.12, 0.12, 0.12),
      opacity: signInfo.opacity,
    });
    cursorY -= index === 0 ? lineHeight + 1 : lineHeight;
  });
};

const updateStatus = (message) => {
  statusEl.textContent = message;
};

const resetDownloadLink = () => {
  if (currentBlobUrl) {
    URL.revokeObjectURL(currentBlobUrl);
    currentBlobUrl = null;
  }
  downloadLink.href = "#";
  downloadLink.classList.add("disabled");
  downloadLink.setAttribute("aria-disabled", "true");
  scheduleLivePreviewUpdate();
};

const setDownloadLink = (url, filename) => {
  if (currentBlobUrl) {
    URL.revokeObjectURL(currentBlobUrl);
  }
  currentBlobUrl = url;
  downloadLink.href = url;
  downloadLink.download = filename;
  downloadLink.classList.remove("disabled");
  downloadLink.setAttribute("aria-disabled", "false");
};

const revokeLivePreviewUrls = () => {
  if (livePreviewBeforeUrl) {
    URL.revokeObjectURL(livePreviewBeforeUrl);
    livePreviewBeforeUrl = null;
  }
  if (livePreviewAfterUrl) {
    URL.revokeObjectURL(livePreviewAfterUrl);
    livePreviewAfterUrl = null;
  }
};

const showLivePreviewEmpty = (message) => {
  if (livePreviewMeta) {
    livePreviewMeta.textContent = message || "Add a file to begin live preview.";
  }
  if (livePreviewEmpty) {
    livePreviewEmpty.textContent = message || "Add a file to begin live preview.";
    livePreviewEmpty.hidden = false;
  }
  if (livePreviewGrid) {
    livePreviewGrid.hidden = true;
  }
};

const updateLivePreviewSourceOptions = () => {
  if (!livePreviewSourceSelect) return;
  const previousValue = livePreviewSourceSelect.value;
  livePreviewSourceSelect.innerHTML = "";

  const autoOption = document.createElement("option");
  autoOption.value = "auto";
  autoOption.textContent = "Use first file in order";
  livePreviewSourceSelect.appendChild(autoOption);

  queue.forEach((item, index) => {
    const option = document.createElement("option");
    option.value = String(item.id);
    option.textContent = `${index + 1}. ${item.file.name}`;
    livePreviewSourceSelect.appendChild(option);
  });

  const hasPrevious = Array.from(livePreviewSourceSelect.options).some(
    (option) => option.value === previousValue
  );
  livePreviewSourceSelect.value = hasPrevious ? previousValue : "auto";
  livePreviewSourceSelect.disabled = queue.length === 0;
};

const getLivePreviewItem = () => {
  if (!queue.length) return null;
  if (!livePreviewSourceSelect || livePreviewSourceSelect.value === "auto") {
    return queue[0];
  }
  const selectedId = Number(livePreviewSourceSelect.value);
  const selected = queue.find((item) => item.id === selectedId);
  return selected || queue[0];
};

const buildPreviewPdfForItem = async (item, withSettings) => {
  if (!item) return null;

  const previewPdf = await PDFDocument.create();
  const watermarkText = withSettings ? getWatermarkText() : null;
  const footerText = withSettings ? getFooterText() : null;
  const signInfo = withSettings ? getSignSettings() : { hasContent: false };
  const watermarkSettings = getWatermarkSettings();
  const footerSettings = getFooterSettings();
  const watermarkFont = watermarkText ? await previewPdf.embedFont(StandardFonts.HelveticaBold) : null;
  const footerFont = footerText ? await previewPdf.embedFont(StandardFonts.Helvetica) : null;
  const signBaseFont = signInfo.hasContent ? await previewPdf.embedFont(StandardFonts.Helvetica) : null;
  const signNameFont = signInfo.hasContent ? await previewPdf.embedFont(StandardFonts.HelveticaBold) : null;
  let signImage = null;
  if (signInfo.hasContent && signInfo.signatureDataUrl) {
    try {
      const bytes = dataUrlToBytes(signInfo.signatureDataUrl);
      if (bytes) {
        signImage = signInfo.signatureDataUrl.startsWith("data:image/jpeg")
          ? await previewPdf.embedJpg(bytes)
          : await previewPdf.embedPng(bytes);
      }
    } catch (error) {
      console.error(error);
    }
  }
  const pageKey = `${item.id}:0`;

  const applySettingsToPage = (page) => {
    if (!withSettings) return;
    if (shouldApplyWatermark(pageKey)) {
      applyWatermark(page, watermarkText, watermarkFont, watermarkSettings, watermarkBox);
    }
    if (shouldApplyFooter(pageKey)) {
      applyFooter(page, footerText, footerFont, footerSettings);
    }
    if (shouldApplySign(pageKey)) {
      applyFillAndSign(
        page,
        signInfo,
        { baseFont: signBaseFont, nameFont: signNameFont },
        signImage
      );
    }
  };

  const bytes = await item.file.arrayBuffer();
  if (item.kind === "pdf") {
    const sourcePdf = await PDFDocument.load(bytes);
    const firstPage = sourcePdf.getPage(0);
    if (!firstPage) {
      const blank = previewPdf.addPage([PAGE_SIZES.letter.width, PAGE_SIZES.letter.height]);
      applySettingsToPage(blank);
    } else if (getPdfPageFitMode() === "original") {
      const [copiedPage] = await previewPdf.copyPages(sourcePdf, [0]);
      const page = previewPdf.addPage(copiedPage);
      applySettingsToPage(page);
    } else {
      const { mediaBox, trimmedBox } = getPdfPageBoxes(firstPage);
      const fallbackBox = sizeToBox(firstPage.getSize()) || sizeToBox(PAGE_SIZES.letter);
      const targetBox = trimmedBox || mediaBox || fallbackBox;
      const boundingBox = toBoundingBox(targetBox);
      const [embeddedPage] = await previewPdf.embedPages([firstPage], [boundingBox]);
      const page = previewPdf.addPage([targetBox.width, targetBox.height]);
      page.drawPage(embeddedPage, {
        x: 0,
        y: 0,
        width: targetBox.width,
        height: targetBox.height,
      });
      applySettingsToPage(page);
    }
  } else if (item.kind === "jpeg" || item.kind === "png") {
    const image = item.kind === "jpeg" ? await previewPdf.embedJpg(bytes) : await previewPdf.embedPng(bytes);
    const { width: imageWidth, height: imageHeight } = image.scale(1);
    const imagePageSize = await resolveImagePageSize();
    const pageWidth = imagePageSize?.width ?? imageWidth;
    const pageHeight = imagePageSize?.height ?? imageHeight;
    const placement = fitImageToPage(imageWidth, imageHeight, pageWidth, pageHeight);
    const page = previewPdf.addPage([pageWidth, pageHeight]);
    page.drawImage(image, placement);
    applySettingsToPage(page);
  }

  const pdfBytes = await previewPdf.save();
  return new Blob([pdfBytes], { type: "application/pdf" });
};

const updateLivePreview = async () => {
  if (
    !livePreviewGrid ||
    !livePreviewEmpty ||
    !livePreviewMeta ||
    !livePreviewBeforeFrame ||
    !livePreviewAfterFrame
  ) {
    return;
  }

  const item = getLivePreviewItem();
  if (!item) {
    revokeLivePreviewUrls();
    livePreviewBeforeFrame.removeAttribute("src");
    livePreviewAfterFrame.removeAttribute("src");
    showLivePreviewEmpty("Add a file to begin live preview.");
    return;
  }

  const token = ++livePreviewToken;
  livePreviewMeta.textContent = `Previewing page 1: ${item.file.name}`;
  livePreviewEmpty.textContent = "Refreshing preview...";
  livePreviewEmpty.hidden = false;
  livePreviewGrid.hidden = true;

  try {
    const [beforeBlob, afterBlob] = await Promise.all([
      buildPreviewPdfForItem(item, false),
      buildPreviewPdfForItem(item, true),
    ]);
    if (token !== livePreviewToken) {
      return;
    }

    const nextBeforeUrl = URL.createObjectURL(beforeBlob);
    const nextAfterUrl = URL.createObjectURL(afterBlob);

    if (livePreviewBeforeUrl) {
      URL.revokeObjectURL(livePreviewBeforeUrl);
    }
    if (livePreviewAfterUrl) {
      URL.revokeObjectURL(livePreviewAfterUrl);
    }

    livePreviewBeforeUrl = nextBeforeUrl;
    livePreviewAfterUrl = nextAfterUrl;
    livePreviewBeforeFrame.src = livePreviewBeforeUrl;
    livePreviewAfterFrame.src = livePreviewAfterUrl;
    livePreviewEmpty.hidden = true;
    livePreviewGrid.hidden = false;
  } catch (error) {
    console.error(error);
    showLivePreviewEmpty("Live preview is unavailable for this file right now.");
  }
};

const scheduleLivePreviewUpdate = () => {
  if (livePreviewDebounceTimer) {
    window.clearTimeout(livePreviewDebounceTimer);
  }
  livePreviewDebounceTimer = window.setTimeout(() => {
    livePreviewDebounceTimer = null;
    updateLivePreview();
  }, 160);
};

const updateUI = () => {
  fileList.innerHTML = "";

  queue.forEach((item, index) => {
    const li = document.createElement("li");
    li.className = "file-item";
    li.dataset.id = String(item.id);
    li.draggable = true;
    li.title = "Drag to change order";

    const order = document.createElement("div");
    order.className = "file-order";
    order.textContent = String(index + 1);

    const meta = document.createElement("div");
    meta.className = "file-meta";

    const name = document.createElement("div");
    name.className = "file-name";
    name.textContent = item.file.name;

    const size = document.createElement("div");
    size.className = "file-size";
    const kindLabel =
      item.kind === "pdf" ? "PDF" : item.kind === "jpeg" ? "JPG" : item.kind === "png" ? "PNG" : "FILE";
    size.textContent = `${kindLabel} - ${formatSize(item.file.size)}`;

    meta.appendChild(name);
    meta.appendChild(size);

    const actions = document.createElement("div");
    actions.className = "file-actions";

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "Remove";
    removeBtn.dataset.id = String(item.id);
    actions.appendChild(removeBtn);

    li.appendChild(order);
    li.appendChild(meta);
    li.appendChild(actions);
    fileList.appendChild(li);
  });

  const hasFiles = queue.length > 0;
  emptyState.style.display = hasFiles ? "none" : "block";
  mergeBtn.disabled = !hasFiles;
  clearBtn.disabled = !hasFiles;
  if (!hasFiles) {
    resetDownloadLink();
    updateStatus("Ready. Add files to begin.");
  }

  updateLivePreviewSourceOptions();
  updatePageSelectors();
  scheduleLivePreviewUpdate();
};

const buildPageEntries = () => {
  const entries = [];
  let globalIndex = 1;

  queue.forEach((item) => {
    if (item.pageCount == null) {
      entries.push({
        loading: true,
        label: `Preparing pages for ${item.file.name}...`,
      });
      return;
    }

    const count = Math.max(1, item.pageCount);
    for (let i = 0; i < count; i += 1) {
      entries.push({
        key: `${item.id}:${i}`,
        title: `Page ${globalIndex}`,
        meta: `${item.file.name} - source page ${i + 1}`,
      });
      globalIndex += 1;
    }
  });

  return entries;
};

const renderPageSelector = ({
  listEl,
  emptyEl,
  selectedSet,
  knownSet,
  defaultSelected,
  setKeys,
  selectAllBtn,
  selectNoneBtn,
  entries,
}) => {
  if (!listEl || !emptyEl) return;

  listEl.innerHTML = "";
  const fragment = document.createDocumentFragment();
  const nextKeys = [];
  let hasEntries = false;

  entries.forEach((entry) => {
    if (entry.loading) {
      const loading = document.createElement("div");
      loading.className = "footer-page-loading";
      loading.textContent = entry.label;
      fragment.appendChild(loading);
      hasEntries = true;
      return;
    }

    const key = entry.key;
    nextKeys.push(key);
    if (!knownSet.has(key)) {
      knownSet.add(key);
      if (defaultSelected) {
        selectedSet.add(key);
      }
    }

    const label = document.createElement("label");
    label.className = "footer-page-item";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = selectedSet.has(key);
    checkbox.dataset.pageKey = key;

    const textWrap = document.createElement("span");
    textWrap.className = "footer-page-text";

    const title = document.createElement("span");
    title.className = "footer-page-title";
    title.textContent = entry.title;

    const meta = document.createElement("span");
    meta.className = "footer-page-meta";
    meta.textContent = entry.meta;

    textWrap.appendChild(title);
    textWrap.appendChild(meta);

    label.appendChild(checkbox);
    label.appendChild(textWrap);
    fragment.appendChild(label);
    hasEntries = true;
  });

  const nextKeySet = new Set(nextKeys);
  selectedSet.forEach((key) => {
    if (!nextKeySet.has(key)) {
      selectedSet.delete(key);
    }
  });
  knownSet.forEach((key) => {
    if (!nextKeySet.has(key)) {
      knownSet.delete(key);
    }
  });

  setKeys(nextKeys);
  listEl.appendChild(fragment);
  listEl.style.display = hasEntries ? "grid" : "none";
  emptyEl.style.display = hasEntries ? "none" : "block";

  const hasSelectablePages = nextKeys.length > 0;
  if (selectAllBtn) {
    selectAllBtn.disabled = !hasSelectablePages;
  }
  if (selectNoneBtn) {
    selectNoneBtn.disabled = !hasSelectablePages;
  }
};

const updatePageSelectors = () => {
  const entries = buildPageEntries();
  renderPageSelector({
    listEl: footerPageList,
    emptyEl: footerPageEmpty,
    selectedSet: selectedFooterPages,
    knownSet: knownFooterPageKeys,
    defaultSelected: footerDefaultSelected,
    setKeys: (keys) => {
      footerPageKeys = keys;
    },
    selectAllBtn: footerSelectAllBtn,
    selectNoneBtn: footerSelectNoneBtn,
    entries,
  });

  renderPageSelector({
    listEl: watermarkPageList,
    emptyEl: watermarkPageEmpty,
    selectedSet: selectedWatermarkPages,
    knownSet: knownWatermarkPageKeys,
    defaultSelected: watermarkDefaultSelected,
    setKeys: (keys) => {
      watermarkPageKeys = keys;
    },
    selectAllBtn: watermarkSelectAllBtn,
    selectNoneBtn: watermarkSelectNoneBtn,
    entries,
  });

  renderPageSelector({
    listEl: signPageList,
    emptyEl: signPageEmpty,
    selectedSet: selectedSignPages,
    knownSet: knownSignPageKeys,
    defaultSelected: signDefaultSelected,
    setKeys: (keys) => {
      signPageKeys = keys;
    },
    selectAllBtn: signSelectAllBtn,
    selectNoneBtn: signSelectNoneBtn,
    entries,
  });

  scheduleWatermarkPreviewUpdate();
  scheduleLivePreviewUpdate();
};

const shouldApplyFooter = (pageKey) =>
  selectedFooterPages.has(pageKey) || (!knownFooterPageKeys.has(pageKey) && footerDefaultSelected);

const shouldApplyWatermark = (pageKey) =>
  selectedWatermarkPages.has(pageKey) ||
  (!knownWatermarkPageKeys.has(pageKey) && watermarkDefaultSelected);

const shouldApplySign = (pageKey) =>
  selectedSignPages.has(pageKey) || (!knownSignPageKeys.has(pageKey) && signDefaultSelected);

const loadPdfPageCount = async (item) => {
  try {
    const bytes = await item.file.arrayBuffer();
    const pdfDoc = await PDFDocument.load(bytes);
    const pageCount = pdfDoc.getPageCount();
    const firstPage = pdfDoc.getPage(0);
    const { mediaBox, trimmedBox } = getPdfPageBoxes(firstPage);
    const size = mediaBox ?? firstPage.getSize();
    const queueItem = queue.find((entry) => entry.id === item.id);
    if (!queueItem) return;
    queueItem.pageCount = pageCount;
    queueItem.pageSize = { width: size.width, height: size.height };
    if (trimmedBox) {
      queueItem.pageTrimSize = { width: trimmedBox.width, height: trimmedBox.height };
    } else {
      queueItem.pageTrimSize = null;
    }
  } catch (error) {
    console.error(error);
    const queueItem = queue.find((entry) => entry.id === item.id);
    if (queueItem) {
      queueItem.pageCount = 1;
      queueItem.pageSize = null;
      queueItem.pageTrimSize = null;
    }
  } finally {
    updatePageSelectors();
  }
};

const addFiles = (files) => {
  const incoming = Array.from(files || []);
  if (!incoming.length) return;

  let accepted = 0;
  let rejected = 0;

  incoming.forEach((file) => {
    const kind = getFileKind(file);
    if (!kind) {
      rejected += 1;
      return;
    }
    const item = {
      id: nextId++,
      file,
      kind,
      pageCount: kind === "pdf" ? null : 1,
      pageSize: null,
      pageTrimSize: null,
    };
    queue.push(item);
    if (kind === "pdf") {
      loadPdfPageCount(item);
    }
    accepted += 1;
  });

  if (accepted > 0) {
    resetDownloadLink();
    updateUI();
  }

  if (accepted === 0 && rejected > 0) {
    updateStatus("Please add only PDF, JPG, or PNG files.");
  } else if (accepted > 0) {
    const suffix = rejected > 0 ? ` Skipped ${rejected} file(s) we cannot use.` : "";
    updateStatus(`Added ${accepted} file(s).${suffix}`);
  }
};

const handleDrop = (event) => {
  if (event.dataTransfer?.files?.length) {
    addFiles(event.dataTransfer.files);
  }
};

const stitchPdfs = async () => {
  if (!queue.length) return;

  mergeBtn.disabled = true;
  mergeBtn.classList.add("is-loading");
  updateStatus(`Building your PDF from ${queue.length} file(s)...`);

  try {
    const mergedPdf = await PDFDocument.create();
    const hasImages = queue.some((item) => item.kind === "jpeg" || item.kind === "png");
    const imagePageSize = hasImages ? await resolveImagePageSize() : null;
    const pdfPageFitMode = getPdfPageFitMode();
    const watermarkText = getWatermarkText();
    const watermarkSettings = getWatermarkSettings();
    const watermarkFont = watermarkText ? await mergedPdf.embedFont(StandardFonts.HelveticaBold) : null;
    const footerText = getFooterText();
    const footerSettings = getFooterSettings();
    const footerFont = footerText ? await mergedPdf.embedFont(StandardFonts.Helvetica) : null;
    const signInfo = getSignSettings();
    const signBaseFont = signInfo.hasContent ? await mergedPdf.embedFont(StandardFonts.Helvetica) : null;
    const signNameFont = signInfo.hasContent ? await mergedPdf.embedFont(StandardFonts.HelveticaBold) : null;
    let signImage = null;
    if (signInfo.hasContent && signInfo.signatureDataUrl) {
      try {
        const signBytes = dataUrlToBytes(signInfo.signatureDataUrl);
        if (signBytes) {
          signImage = signInfo.signatureDataUrl.startsWith("data:image/jpeg")
            ? await mergedPdf.embedJpg(signBytes)
            : await mergedPdf.embedPng(signBytes);
        }
      } catch (error) {
        console.error(error);
      }
    }

    for (const item of queue) {
      const bytes = await item.file.arrayBuffer();
      if (item.kind === "pdf") {
        const sourcePdf = await PDFDocument.load(bytes);
        const pageIndices = sourcePdf.getPageIndices();
        if (pdfPageFitMode === "original") {
          const copiedPages = await mergedPdf.copyPages(sourcePdf, pageIndices);
          copiedPages.forEach((page, pageIndex) => {
            const addedPage = mergedPdf.addPage(page);
            const pageKey = `${item.id}:${pageIndex}`;
            if (shouldApplyWatermark(pageKey)) {
              applyWatermark(addedPage, watermarkText, watermarkFont, watermarkSettings, watermarkBox);
            }
            if (shouldApplyFooter(pageKey)) {
              applyFooter(addedPage, footerText, footerFont, footerSettings);
            }
            if (shouldApplySign(pageKey)) {
              applyFillAndSign(
                addedPage,
                signInfo,
                { baseFont: signBaseFont, nameFont: signNameFont },
                signImage
              );
            }
          });
        } else {
          for (const pageIndex of pageIndices) {
            const sourcePage = sourcePdf.getPage(pageIndex);
            const { mediaBox, trimmedBox } = getPdfPageBoxes(sourcePage);
            const targetBox = pdfPageFitMode === "trim" ? trimmedBox : mediaBox;
            const fallbackBox = mediaBox ?? sizeToBox(sourcePage.getSize());
            const finalBox = targetBox || fallbackBox;
            const boundingBox = toBoundingBox(finalBox);
            const [embeddedPage] = await mergedPdf.embedPages([sourcePage], [boundingBox]);
            const addedPage = mergedPdf.addPage([finalBox.width, finalBox.height]);
            addedPage.drawPage(embeddedPage, {
              x: 0,
              y: 0,
              width: finalBox.width,
              height: finalBox.height,
            });
            const pageKey = `${item.id}:${pageIndex}`;
            if (shouldApplyWatermark(pageKey)) {
              applyWatermark(addedPage, watermarkText, watermarkFont, watermarkSettings, watermarkBox);
            }
            if (shouldApplyFooter(pageKey)) {
              applyFooter(addedPage, footerText, footerFont, footerSettings);
            }
            if (shouldApplySign(pageKey)) {
              applyFillAndSign(
                addedPage,
                signInfo,
                { baseFont: signBaseFont, nameFont: signNameFont },
                signImage
              );
            }
          }
        }
      } else if (item.kind === "jpeg" || item.kind === "png") {
        const image =
          item.kind === "jpeg" ? await mergedPdf.embedJpg(bytes) : await mergedPdf.embedPng(bytes);
        const { width: imageWidth, height: imageHeight } = image.scale(1);
        const pageWidth = imagePageSize?.width ?? imageWidth;
        const pageHeight = imagePageSize?.height ?? imageHeight;
        const placement = fitImageToPage(imageWidth, imageHeight, pageWidth, pageHeight);
        const page = mergedPdf.addPage([pageWidth, pageHeight]);
        page.drawImage(image, placement);
        const pageKey = `${item.id}:0`;
        if (shouldApplyWatermark(pageKey)) {
          applyWatermark(page, watermarkText, watermarkFont, watermarkSettings, watermarkBox);
        }
        if (shouldApplyFooter(pageKey)) {
          applyFooter(page, footerText, footerFont, footerSettings);
        }
        if (shouldApplySign(pageKey)) {
          applyFillAndSign(
            page,
            signInfo,
            { baseFont: signBaseFont, nameFont: signNameFont },
            signImage
          );
        }
      }
    }

    const mergedBytes = await mergedPdf.save();
    const blob = new Blob([mergedBytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const dateStamp = new Date().toISOString().slice(0, 10);

    setDownloadLink(url, `stitched-${dateStamp}.pdf`);
    updateStatus("Done. Your PDF is ready to download.");
  } catch (error) {
    console.error(error);
    updateStatus("Something went wrong while building your PDF. Please try again.");
  } finally {
    mergeBtn.disabled = queue.length === 0;
    mergeBtn.classList.remove("is-loading");
  }
};

browseBtn.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", (event) => {
  addFiles(event.target.files);
  fileInput.value = "";
});

let dragDepth = 0;
let draggedId = null;

const getDragItem = (event) => {
  const target = event.target;
  if (!(target instanceof Element)) return null;
  return target.closest(".file-item");
};

const clearDropIndicators = () => {
  fileList
    .querySelectorAll(".file-item.drag-over-top, .file-item.drag-over-bottom")
    .forEach((item) => {
      item.classList.remove("drag-over-top", "drag-over-bottom");
    });
};

dropzone.addEventListener("dragenter", (event) => {
  event.preventDefault();
  dragDepth += 1;
  dropzone.classList.add("active");
});

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
});

dropzone.addEventListener("dragleave", (event) => {
  event.preventDefault();
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) {
    dropzone.classList.remove("active");
  }
});

dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dragDepth = 0;
  dropzone.classList.remove("active");
  handleDrop(event);
});

fileList.addEventListener("dragstart", (event) => {
  const item = getDragItem(event);
  if (!item) return;
  draggedId = Number(item.dataset.id);
  item.classList.add("dragging");
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", String(draggedId));
  }
});

fileList.addEventListener("dragover", (event) => {
  if (draggedId === null) return;
  const item = getDragItem(event);
  event.preventDefault();
  clearDropIndicators();
  if (!item) return;
  const rect = item.getBoundingClientRect();
  const insertAfter = event.clientY > rect.top + rect.height / 2;
  item.classList.add(insertAfter ? "drag-over-bottom" : "drag-over-top");
});

fileList.addEventListener("drop", (event) => {
  if (draggedId === null) return;
  const item = getDragItem(event);
  event.preventDefault();
  clearDropIndicators();
  if (!item) {
    const draggedIndex = queue.findIndex((entry) => entry.id === draggedId);
    if (draggedIndex < 0) return;
    const [moved] = queue.splice(draggedIndex, 1);
    queue.push(moved);
    resetDownloadLink();
    updateUI();
    updateStatus("File order saved.");
    draggedId = null;
    return;
  }
  const targetId = Number(item.dataset.id);
  if (targetId === draggedId) return;
  const rect = item.getBoundingClientRect();
  const insertAfter = event.clientY > rect.top + rect.height / 2;
  const draggedIndex = queue.findIndex((entry) => entry.id === draggedId);
  if (draggedIndex < 0) return;
  const [moved] = queue.splice(draggedIndex, 1);
  const targetIndex = queue.findIndex((entry) => entry.id === targetId);
  const insertIndex = insertAfter ? targetIndex + 1 : targetIndex;
  queue.splice(insertIndex, 0, moved);
  resetDownloadLink();
  updateUI();
  updateStatus("File order saved.");
  draggedId = null;
});

fileList.addEventListener("dragend", () => {
  const draggingItem = fileList.querySelector(".file-item.dragging");
  if (draggingItem) {
    draggingItem.classList.remove("dragging");
  }
  clearDropIndicators();
  draggedId = null;
});

fileList.addEventListener("click", (event) => {
  const target = event.target;
  if (target instanceof HTMLButtonElement && target.dataset.id) {
    const id = Number(target.dataset.id);
    const index = queue.findIndex((item) => item.id === id);
    if (index >= 0) {
      queue.splice(index, 1);
      resetDownloadLink();
      updateUI();
    }
  }
});

clearBtn.addEventListener("click", () => {
  queue.length = 0;
  resetDownloadLink();
  updateUI();
  updateStatus("All files removed.");
});

mergeBtn.addEventListener("click", stitchPdfs);

if (imagePageSizeSelect) {
  imagePageSizeSelect.addEventListener("change", () => {
    resetDownloadLink();
    updateStatus("Photo page size saved.");
    scheduleWatermarkPreviewUpdate();
  });
}

if (pdfPageFitSelect) {
  pdfPageFitSelect.addEventListener("change", () => {
    resetDownloadLink();
    updateStatus("PDF page framing saved.");
    scheduleWatermarkPreviewUpdate();
  });
}

if (watermarkInput) {
  watermarkInput.addEventListener("input", () => {
    resetDownloadLink();
    updateStatus(watermarkInput.value.trim() ? "Watermark text saved." : "Watermark text cleared.");
    scheduleWatermarkPreviewUpdate();
  });
}

if (watermarkOpacityInput) {
  watermarkOpacityInput.addEventListener("input", () => {
    resetDownloadLink();
    scheduleWatermarkPreviewUpdate();
  });
  watermarkOpacityInput.addEventListener("change", () => {
    updateStatus("Watermark style saved.");
  });
}

if (watermarkToneInput) {
  watermarkToneInput.addEventListener("input", () => {
    resetDownloadLink();
    scheduleWatermarkPreviewUpdate();
  });
  watermarkToneInput.addEventListener("change", () => {
    updateStatus("Watermark style saved.");
  });
}

if (watermarkSizeInput) {
  watermarkSizeInput.addEventListener("input", () => {
    resetDownloadLink();
    scheduleWatermarkPreviewUpdate();
  });
  watermarkSizeInput.addEventListener("change", () => {
    updateStatus("Watermark size saved.");
  });
}

if (watermarkOrientationSelect) {
  watermarkOrientationSelect.addEventListener("change", () => {
    resetDownloadLink();
    scheduleWatermarkPreviewUpdate();
    updateStatus("Watermark angle saved.");
  });
}

if (watermarkPageList) {
  watermarkPageList.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (target.type !== "checkbox" || !target.dataset.pageKey) return;
    const key = target.dataset.pageKey;
    if (target.checked) {
      selectedWatermarkPages.add(key);
    } else {
      selectedWatermarkPages.delete(key);
    }
    resetDownloadLink();
    updateStatus("Watermark page choices saved.");
  });
}

if (watermarkSelectAllBtn) {
  watermarkSelectAllBtn.addEventListener("click", () => {
    watermarkDefaultSelected = true;
    watermarkPageKeys.forEach((key) => {
      selectedWatermarkPages.add(key);
    });
    updatePageSelectors();
    resetDownloadLink();
    updateStatus("Watermark page choices saved.");
  });
}

if (watermarkSelectNoneBtn) {
  watermarkSelectNoneBtn.addEventListener("click", () => {
    watermarkDefaultSelected = false;
    selectedWatermarkPages.clear();
    updatePageSelectors();
    resetDownloadLink();
    updateStatus("Watermark page choices saved.");
  });
}

if (footerPageList) {
  footerPageList.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (target.type !== "checkbox" || !target.dataset.pageKey) return;
    const key = target.dataset.pageKey;
    if (target.checked) {
      selectedFooterPages.add(key);
    } else {
      selectedFooterPages.delete(key);
    }
    resetDownloadLink();
    updateStatus("Bottom note page choices saved.");
  });
}

if (footerSelectAllBtn) {
  footerSelectAllBtn.addEventListener("click", () => {
    footerDefaultSelected = true;
    footerPageKeys.forEach((key) => {
      selectedFooterPages.add(key);
    });
    updatePageSelectors();
    resetDownloadLink();
    updateStatus("Bottom note page choices saved.");
  });
}

if (footerSelectNoneBtn) {
  footerSelectNoneBtn.addEventListener("click", () => {
    footerDefaultSelected = false;
    selectedFooterPages.clear();
    updatePageSelectors();
    resetDownloadLink();
    updateStatus("Bottom note page choices saved.");
  });
}

if (signNameInput) {
  signNameInput.addEventListener("input", () => {
    resetDownloadLink();
    updateStatus(signNameInput.value.trim() ? "Signer name saved." : "Signer name cleared.");
    scheduleWatermarkPreviewUpdate();
  });
}

if (signDateInput) {
  const handleSignDateChange = () => {
    resetDownloadLink();
    updateStatus(signDateInput.value ? "Signer date saved." : "Signer date cleared.");
    scheduleWatermarkPreviewUpdate();
  };
  signDateInput.addEventListener("input", handleSignDateChange);
  signDateInput.addEventListener("change", handleSignDateChange);
}

if (signNoteInput) {
  signNoteInput.addEventListener("input", () => {
    resetDownloadLink();
    updateStatus(signNoteInput.value.trim() ? "Signer note saved." : "Signer note cleared.");
    scheduleWatermarkPreviewUpdate();
  });
}

if (signPlacementSelect) {
  signPlacementSelect.addEventListener("change", () => {
    resetDownloadLink();
    updateStatus("Fill and sign position saved.");
    scheduleWatermarkPreviewUpdate();
  });
}

if (signOpacityInput) {
  signOpacityInput.addEventListener("input", () => {
    resetDownloadLink();
    scheduleWatermarkPreviewUpdate();
  });
  signOpacityInput.addEventListener("change", () => {
    updateStatus("Fill and sign style saved.");
  });
}

if (signaturePad) {
  signaturePadContext = signaturePad.getContext("2d");
  syncSignaturePadCanvasSize();
  signaturePad.addEventListener("pointerdown", startSignatureDraw);
  signaturePad.addEventListener("pointermove", moveSignatureDraw);
  signaturePad.addEventListener("pointerup", endSignatureDraw);
  signaturePad.addEventListener("pointercancel", endSignatureDraw);
}

if (signatureClearBtn) {
  signatureClearBtn.addEventListener("click", () => {
    clearSignaturePad();
    resetDownloadLink();
    scheduleWatermarkPreviewUpdate();
    updateStatus("Signature cleared.");
  });
}

if (signatureUploadInput) {
  signatureUploadInput.addEventListener("change", () => {
    const file = signatureUploadInput.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      signatureUploadInput.value = "";
      updateStatus("Please upload a signature image (PNG or JPG).");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = typeof reader.result === "string" ? reader.result : null;
      if (!dataUrl) {
        updateStatus("We could not read that signature image. Please try another one.");
        return;
      }
      signatureDataUrl = dataUrl;
      loadSignatureDataUrlToPad(dataUrl);
      resetDownloadLink();
      scheduleWatermarkPreviewUpdate();
      updateStatus("Signature image saved.");
    };
    reader.onerror = () => {
      updateStatus("We could not read that signature image. Please try another one.");
    };
    reader.readAsDataURL(file);
  });
}

if (signPageList) {
  signPageList.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (target.type !== "checkbox" || !target.dataset.pageKey) return;
    const key = target.dataset.pageKey;
    if (target.checked) {
      selectedSignPages.add(key);
    } else {
      selectedSignPages.delete(key);
    }
    resetDownloadLink();
    updateStatus("Fill and sign page choices saved.");
  });
}

if (signSelectAllBtn) {
  signSelectAllBtn.addEventListener("click", () => {
    signDefaultSelected = true;
    signPageKeys.forEach((key) => {
      selectedSignPages.add(key);
    });
    updatePageSelectors();
    resetDownloadLink();
    updateStatus("Fill and sign page choices saved.");
  });
}

if (signSelectNoneBtn) {
  signSelectNoneBtn.addEventListener("click", () => {
    signDefaultSelected = false;
    selectedSignPages.clear();
    updatePageSelectors();
    resetDownloadLink();
    updateStatus("Fill and sign page choices saved.");
  });
}

previewFrames.forEach((frame) => {
  frame.addEventListener("pointerdown", (event) => {
    startWatermarkDrag(event, "move");
  });
});

previewHandles.forEach((handle) => {
  handle.addEventListener("pointerdown", (event) => {
    event.stopPropagation();
    startWatermarkDrag(event, "resize");
  });
});

if (footerInput) {
  footerInput.addEventListener("input", () => {
    resetDownloadLink();
    updateStatus(footerInput.value.trim() ? "Bottom note text saved." : "Bottom note text cleared.");
    scheduleWatermarkPreviewUpdate();
  });
}

if (footerOpacityInput) {
  footerOpacityInput.addEventListener("input", () => {
    resetDownloadLink();
    scheduleWatermarkPreviewUpdate();
  });
  footerOpacityInput.addEventListener("change", () => {
    updateStatus("Bottom note style saved.");
  });
}

if (footerToneInput) {
  footerToneInput.addEventListener("input", () => {
    resetDownloadLink();
    scheduleWatermarkPreviewUpdate();
  });
  footerToneInput.addEventListener("change", () => {
    updateStatus("Bottom note style saved.");
  });
}

if (previewOrientationSelect) {
  previewOrientationSelect.addEventListener("change", () => {
    scheduleWatermarkPreviewUpdate();
    updateStatus("Preview page shape saved.");
  });
}

if (livePreviewSourceSelect) {
  livePreviewSourceSelect.addEventListener("change", () => {
    scheduleLivePreviewUpdate();
  });
}

downloadLink.addEventListener("click", (event) => {
  if (downloadLink.classList.contains("disabled")) {
    event.preventDefault();
  }
});

window.addEventListener("beforeunload", () => {
  if (currentBlobUrl) {
    URL.revokeObjectURL(currentBlobUrl);
  }
  revokeLivePreviewUrls();
});

window.addEventListener("resize", () => {
  syncSignaturePadCanvasSize();
  scheduleWatermarkPreviewUpdate();
  scheduleLivePreviewUpdate();
});

updateUI();
updateWatermarkPreview();
updateLivePreview();
