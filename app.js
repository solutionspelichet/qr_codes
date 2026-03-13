
const CONFIG = window.APP_CONFIG || {};

const state = {
  presets: [],
  rows: [],
  headers: [],
  tokenClient: null,
  accessToken: null,
  generatedFiles: [],
  mainArtifact: null,
  objectUrls: []
};

function $(id) {
  return document.getElementById(id);
}

function checkedValues(selector) {
  return [...document.querySelectorAll(selector)]
    .filter(el => el.checked)
    .map(el => el.value);
}

function setStatus(message, type = "info", loading = false) {
  const box = $("status-box");
  box.className = "rounded-xl p-4 text-sm";
  box.classList.remove("hidden", "status-info", "status-success", "status-error");
  box.classList.add(type === "success" ? "status-success" : type === "error" ? "status-error" : "status-info");
  box.innerHTML = loading
    ? '<div class="flex items-center gap-2"><span class="spinner"></span><span>' + message + '</span></div>'
    : message;
  box.classList.remove("hidden");
}

function clearStatus() {
  $("status-box").classList.add("hidden");
  $("status-box").innerHTML = "";
}

function mmToPx(mm, dpi) {
  return Math.max(1, Math.round((Number(mm) / 25.4) * Number(dpi)));
}

function sanitizeCode39(value) {
  const allowed = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 -.*/+$%";
  return String(value || "")
    .toUpperCase()
    .split("")
    .filter(c => allowed.includes(c))
    .join("");
}

function safeFilename(value) {
  return String(value || "")
    .trim()
    .replace(/[^\w\-.]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^[_\.]+|[_\.]+$/g, "")
    .slice(0, 80) || "label";
}

function clearObjectUrls() {
  state.objectUrls.forEach(url => URL.revokeObjectURL(url));
  state.objectUrls = [];
}

function blobUrl(blob) {
  const url = URL.createObjectURL(blob);
  state.objectUrls.push(url);
  return url;
}

async function loadPresets() {
  const res = await fetch("./presets.json");
  state.presets = await res.json();
  const select = $("preset-select");
  select.innerHTML = "";
  state.presets.forEach(p => {
    const opt = document.createElement("option");
    opt.value = p.id;
    opt.textContent = p.name;
    select.appendChild(opt);
  });
}

function applyPreset(id) {
  const p = state.presets.find(x => x.id === id);
  if (!p || p.id === "custom") return;
  $("label-width-mm").value = p.label_width_mm;
  $("label-height-mm").value = p.label_height_mm;
  $("spacing-x-mm").value = p.spacing_x_mm;
  $("spacing-y-mm").value = p.spacing_y_mm;
  $("margin-top-mm").value = p.margin_top_mm;
  $("margin-left-mm").value = p.margin_left_mm;
  $("columns-per-page").value = p.columns_per_page;
  $("rows-per-page").value = p.rows_per_page;
  renderPreview();
}

async function parseExcelFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheet];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const headers = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" })[0] || [];

  state.rows = rows;
  state.headers = headers;

  const select = $("column-name");
  select.innerHTML = "";
  headers.forEach(h => {
    const opt = document.createElement("option");
    opt.value = String(h);
    opt.textContent = String(h);
    select.appendChild(opt);
  });

  const preferred = headers.find(h => String(h).trim().toLowerCase() === "code") || headers[0];
  if (preferred) {
    select.value = preferred;
  }

  const codes = getCodes().codes;
  $("preview-code").value = codes[0] || "";
  $("excel-summary").textContent = rows.length + " ligne(s) detectee(s) - feuille: " + firstSheet;
  renderPreview();
}

function getCodes() {
  const column = $("column-name").value;
  const codes = state.rows.map(r => String(r[column] ?? "").trim()).filter(Boolean);
  return { column, codes };
}

function collectOptions() {
  return {
    code_types: checkedValues('input[data-role="code-type"]'),
    output_formats: checkedValues('input[data-role="format"]'),
    label_width_mm: Number($("label-width-mm").value),
    label_height_mm: Number($("label-height-mm").value),
    spacing_x_mm: Number($("spacing-x-mm").value),
    spacing_y_mm: Number($("spacing-y-mm").value),
    margin_top_mm: Number($("margin-top-mm").value),
    margin_left_mm: Number($("margin-left-mm").value),
    columns_per_page: Number($("columns-per-page").value),
    rows_per_page: Number($("rows-per-page").value),
    barcode_height: Number($("barcode-height").value),
    font_size: Number($("font-size").value),
    text_margin_mm: Number($("text-margin-mm").value),
    dpi: Number($("dpi").value || CONFIG.DEFAULT_DPI || 200),
    colors: {
      code: $("code-color").value,
      text: $("text-color").value,
      background: $("background-color").value
    }
  };
}

function buildTemplateExcel() {
  const ws = XLSX.utils.aoa_to_sheet([
    ["Code"],
    ["001234"],
    ["ABC123"],
    ["LOT-A-01"]
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Etiquettes");
  XLSX.writeFile(wb, "modele_etiquettes.xlsx");
}

function renderPreview() {
  const box = $("preview-label");
  box.innerHTML = "";

  const value = $("preview-code").value.trim() || "ABC123";
  const types = checkedValues('input[data-role="code-type"]');
  const codeColor = $("code-color").value;
  const backgroundColor = $("background-color").value;

  if (!types.length) {
    box.innerHTML = '<div class="text-sm text-slate-500">Aucun type selectionne</div>';
    return;
  }

  types.forEach(type => {
    const wrapper = document.createElement("div");
    wrapper.className = "text-center";

    const title = document.createElement("div");
    title.className = "text-xs font-semibold mb-2 text-slate-500";
    title.textContent = type;
    wrapper.appendChild(title);

    if (type === "QR") {
      const canvas = document.createElement("canvas");
      new QRious({
        element: canvas,
        value: value,
        size: 100,
        foreground: codeColor,
        background: backgroundColor
      });
      wrapper.appendChild(canvas);
    } else {
      const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      JsBarcode(svg, type === "Code39" ? (sanitizeCode39(value) || "ABC123") : value, {
        format: type === "Code39" ? "CODE39" : "CODE128",
        lineColor: codeColor,
        background: backgroundColor,
        displayValue: true,
        margin: 0,
        width: 1.5,
        height: 40
      });
      wrapper.appendChild(svg);
    }

    box.appendChild(wrapper);
  });
}

function fitFont(ctx, text, maxWidth, startSize, minSize) {
  let size = startSize;
  while (size >= minSize) {
    ctx.font = size + "px Arial";
    if (ctx.measureText(text).width <= maxWidth) return size;
    size--;
  }
  return minSize;
}

function drawContain(ctx, img, x, y, maxW, maxH) {
  const ratio = Math.min(maxW / img.width, maxH / img.height);
  const w = img.width * ratio;
  const h = img.height * ratio;
  ctx.drawImage(img, x + (maxW - w) / 2, y + (maxH - h) / 2, w, h);
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Image temporaire impossible a charger"));
    img.src = src;
  });
}

function canvasToBlob(canvas, type = "image/png") {
  return new Promise(resolve => canvas.toBlob(resolve, type));
}

async function barcodeCanvas(code, type, options, width, height) {
  const value = type === "Code39" ? (sanitizeCode39(code) || "EMPTY") : String(code);
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");

  JsBarcode(svg, value, {
    format: type === "Code39" ? "CODE39" : "CODE128",
    lineColor: options.colors.code,
    background: options.colors.background,
    displayValue: false,
    margin: 0,
    width: 2,
    height: Math.max(30, height - 4)
  });

  const svgText = new XMLSerializer().serializeToString(svg);
  const svgUrl = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgText);
  const img = await loadImage(svgUrl);

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = options.colors.background;
  ctx.fillRect(0, 0, width, height);
  drawContain(ctx, img, 0, 0, width, height);
  return canvas;
}

function qrCanvas(code, options, size) {
  const canvas = document.createElement("canvas");
  canvas.width = size;
  canvas.height = size;
  new QRious({
    element: canvas,
    value: String(code),
    size: size,
    foreground: options.colors.code,
    background: options.colors.background
  });
  return canvas;
}

async function createLabelCanvas(code, type, options) {
  const dpi = options.dpi;
  const width = mmToPx(options.label_width_mm, dpi);
  const height = mmToPx(options.label_height_mm, dpi);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext("2d");
  ctx.fillStyle = options.colors.background;
  ctx.fillRect(0, 0, width, height);

  const pad = Math.max(12, Math.round(width * 0.04));
  const textMargin = mmToPx(options.text_margin_mm, dpi);
  const graphicW = width - pad * 2;
  const textBase = Math.max(12, Math.round((options.font_size * 96) / 72));
  const graphicH = height - pad * 2 - textBase - textMargin - 8;

  if (type === "QR") {
    const size = Math.max(20, Math.min(graphicW, graphicH));
    const qr = qrCanvas(code, options, size);
    ctx.drawImage(qr, (width - size) / 2, pad);
  } else {
    const bc = await barcodeCanvas(code, type, options, graphicW, graphicH);
    ctx.drawImage(bc, pad, pad, graphicW, graphicH);
  }

  const fontPx = fitFont(ctx, String(code), graphicW, textBase, 10);
  ctx.font = fontPx + "px Arial";
  ctx.fillStyle = options.colors.text;
  ctx.textAlign = "center";
  ctx.textBaseline = "top";
  ctx.fillText(String(code), width / 2, pad + graphicH + Math.max(4, textMargin / 2));

  return canvas;
}

async function buildPdfBlob(canvases, options) {
  const jsPDF = window.jspdf.jsPDF;
  const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4", compress: true });

  const cols = options.columns_per_page;
  const rows = options.rows_per_page;
  const labelW = options.label_width_mm;
  const labelH = options.label_height_mm;
  const spacingX = options.spacing_x_mm;
  const spacingY = options.spacing_y_mm;
  const marginTop = options.margin_top_mm;
  const marginLeft = options.margin_left_mm;
  const perPage = cols * rows;

  canvases.forEach((canvas, index) => {
    if (index > 0 && index % perPage === 0) pdf.addPage();

    const i = index % perPage;
    const row = Math.floor(i / cols);
    const col = i % cols;
    const x = marginLeft + col * (labelW + spacingX);
    const y = marginTop + row * (labelH + spacingY);

    pdf.addImage(canvas.toDataURL("image/png"), "PNG", x, y, labelW, labelH);
  });

  return pdf.output("blob");
}

function showResults(files, mainArtifact) {
  state.generatedFiles = files;
  state.mainArtifact = mainArtifact;

  const box = $("results-box");
  const list = $("files-list");
  const mainLink = $("main-download-link");

  list.innerHTML = "";
  mainLink.href = mainArtifact.url;
  mainLink.download = mainArtifact.name;

  files.forEach(file => {
    const a = document.createElement("a");
    a.href = file.url;
    a.download = file.name;
    a.textContent = file.name + " (" + Math.round(file.blob.size / 1024) + " Ko)";
    list.appendChild(a);
  });

  box.classList.remove("hidden");
  $("upload-drive-btn").classList.remove("hidden");
}

async function generateAll() {
  clearObjectUrls();
  $("drive-link-box").classList.add("hidden");

  if (!state.rows.length) {
    setStatus("Charge un fichier Excel d'abord.", "error");
    return;
  }

  const options = collectOptions();
  const data = getCodes();
  const codes = data.codes;

  if (!codes.length) {
    setStatus("Aucun code valide.", "error");
    return;
  }

  if (!options.code_types.length) {
    setStatus("Choisis au moins un type de code.", "error");
    return;
  }

  if (!options.output_formats.length) {
    setStatus("Choisis au moins un format.", "error");
    return;
  }

  if (codes.length > Number(CONFIG.MAX_ROWS || 300)) {
    setStatus("Trop de lignes pour le navigateur.", "error");
    return;
  }

  const files = [];
  const pdfGroups = {};
  const btn = $("generate-btn");
  btn.disabled = true;

  try {
    for (const codeType of options.code_types) {
      pdfGroups[codeType] = [];

      for (let i = 0; i < codes.length; i++) {
        setStatus("Generation " + codeType + " " + (i + 1) + "/" + codes.length, "info", true);

        const canvas = await createLabelCanvas(codes[i], codeType, options);

        if (options.output_formats.includes("PNG")) {
          const pngBlob = await canvasToBlob(canvas, "image/png");
          files.push({
            name: "png/" + codeType.toLowerCase() + "/" + String(i + 1).padStart(4, "0") + "_" + safeFilename(codes[i]) + ".png",
            blob: pngBlob,
            url: blobUrl(pngBlob)
          });
        }

        if (options.output_formats.includes("PDF")) {
          pdfGroups[codeType].push(canvas);
        }

        if (i % 5 === 0) {
          await new Promise(resolve => requestAnimationFrame(resolve));
        }
      }
    }

    if (options.output_formats.includes("PDF")) {
      for (const codeType of options.code_types) {
        setStatus("Construction PDF " + codeType, "info", true);
        const pdfBlob = await buildPdfBlob(pdfGroups[codeType], options);
        files.push({
          name: "pdf/labels_" + codeType.toLowerCase() + ".pdf",
          blob: pdfBlob,
          url: blobUrl(pdfBlob)
        });
      }
    }

    let mainArtifact;

    if (files.length === 1) {
      mainArtifact = {
        name: files[0].name.split("/").pop(),
        blob: files[0].blob,
        url: files[0].url
      };
    } else {
      setStatus("Creation ZIP...", "info", true);
      const zip = new JSZip();
      files.forEach(f => zip.file(f.name, f.blob));
      zip.file("manifest.json", JSON.stringify({
        total_codes: codes.length,
        selected_column: data.column,
        options: options
      }, null, 2));
      const zipBlob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
      mainArtifact = {
        name: "etiquettes_" + Date.now() + ".zip",
        blob: zipBlob,
        url: blobUrl(zipBlob)
      };
    }

    showResults(files, mainArtifact);
    setStatus("Generation terminee.", "success");
  } catch (e) {
    console.error(e);
    setStatus("Erreur: " + e.message, "error");
  } finally {
    btn.disabled = false;
  }
}

function initDrive() {
  if (!window.google || !window.google.accounts || !window.google.accounts.oauth2) {
    throw new Error("Module Google non charge");
  }
  if (!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.includes("REMPLACE")) {
    throw new Error("GOOGLE_CLIENT_ID manquant dans config.js");
  }
  if (!state.tokenClient) {
    state.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.GOOGLE_CLIENT_ID,
      scope: "https://www.googleapis.com/auth/drive",
      callback: () => {}
    });
  }
}

function ensureDriveToken() {
  initDrive();
  return new Promise((resolve, reject) => {
    state.tokenClient.callback = (resp) => {
      if (resp.error) {
        reject(new Error(resp.error));
        return;
      }
      state.accessToken = resp.access_token;
      $("drive-status").textContent = "Connecte a Google Drive.";
      resolve(resp.access_token);
    };
    state.tokenClient.requestAccessToken({ prompt: state.accessToken ? "" : "consent" });
  });
}

function buildMultipartBody(metadata, blob, mimeType, boundary) {
  const delimiter = "\r\n--" + boundary + "\r\n";
  const closeDelimiter = "\r\n--" + boundary + "--";
  return new Blob([
    delimiter,
    "Content-Type: application/json; charset=UTF-8\r\n\r\n",
    JSON.stringify(metadata),
    delimiter,
    "Content-Type: " + mimeType + "\r\n\r\n",
    blob,
    closeDelimiter
  ], { type: "multipart/related; boundary=" + boundary });
}

async function makePublic(fileId, token) {
  const res = await fetch("https://www.googleapis.com/drive/v3/files/" + fileId + "/permissions", {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ role: "reader", type: "anyone" })
  });
  if (!res.ok) {
    throw new Error(await res.text());
  }
}

async function uploadToDrive() {
  if (!state.mainArtifact) {
    setStatus("Aucun resultat a envoyer.", "error");
    return;
  }

  const btn = $("upload-drive-btn");
  btn.disabled = true;

  try {
    setStatus("Connexion Google Drive...", "info", true);
    const token = await ensureDriveToken();

    const metadata = { name: state.mainArtifact.name };
    const folderId = $("drive-folder-id").value.trim();
    if (folderId) metadata.parents = [folderId];

    const boundary = "-------314159265358979323846";
    const body = buildMultipartBody(metadata, state.mainArtifact.blob, "application/octet-stream", boundary);

    setStatus("Upload Drive...", "info", true);
    const res = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name", {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + token,
        "Content-Type": "multipart/related; boundary=" + boundary
      },
      body: body
    });

    if (!res.ok) {
      throw new Error(await res.text());
    }

    const payload = await res.json();

    if ($("share-publicly").checked) {
      await makePublic(payload.id, token);
    }

    const url = "https://drive.google.com/file/d/" + payload.id + "/view";
    $("drive-link").href = url;
    $("drive-link").textContent = url;
    $("drive-link-box").classList.remove("hidden");
    setStatus("Fichier envoye sur Google Drive.", "success");
  } catch (e) {
    console.error(e);
    setStatus("Erreur Drive: " + e.message, "error");
  } finally {
    btn.disabled = false;
  }
}

async function connectDrive() {
  try {
    setStatus("Connexion Google Drive...", "info", true);
    await ensureDriveToken();
    setStatus("Connexion Google Drive OK.", "success");
  } catch (e) {
    setStatus("Erreur Drive: " + e.message, "error");
  }
}

function bindEvents() {
  $("download-template-btn").addEventListener("click", buildTemplateExcel);
  $("preset-select").addEventListener("change", e => applyPreset(e.target.value));
  $("excel-file").addEventListener("change", async e => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      await parseExcelFile(file);
      clearStatus();
    } catch (err) {
      setStatus("Lecture Excel impossible: " + err.message, "error");
    }
  });

  $("column-name").addEventListener("change", () => {
    const codes = getCodes().codes;
    $("preview-code").value = codes[0] || "";
    renderPreview();
  });

  [
    "preview-code",
    "code-color",
    "text-color",
    "background-color",
    "label-width-mm",
    "label-height-mm",
    "spacing-x-mm",
    "spacing-y-mm",
    "margin-top-mm",
    "margin-left-mm",
    "columns-per-page",
    "rows-per-page",
    "barcode-height",
    "font-size",
    "text-margin-mm",
    "dpi"
  ].forEach(id => {
    $(id).addEventListener("input", renderPreview);
  });

  document.querySelectorAll('input[data-role="code-type"]').forEach(el => {
    el.addEventListener("change", renderPreview);
  });

  $("generate-btn").addEventListener("click", generateAll);
  $("connect-drive-btn").addEventListener("click", connectDrive);
  $("upload-drive-btn").addEventListener("click", uploadToDrive);
}

async function init() {
  $("drive-folder-id").value = CONFIG.DEFAULT_DRIVE_FOLDER_ID || "";
  $("share-publicly").checked = Boolean(CONFIG.DEFAULT_SHARE_PUBLICLY);
  $("dpi").value = CONFIG.DEFAULT_DPI || 200;
  await loadPresets();
  bindEvents();
  renderPreview();
}

document.addEventListener("DOMContentLoaded", init);
