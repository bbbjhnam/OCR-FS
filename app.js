const TARGET_ITEMS = [
  { key: "assets", label: "자산총계", aliases: ["자산총계", "자 산 총 계"] },
  { key: "liabilities", label: "부채총계", aliases: ["부채총계", "부 채 총 계"] },
  { key: "capital", label: "자본금", aliases: ["자본금", "자 본 금"] },
  { key: "revenue", label: "매출액", aliases: ["매출액", "영업수익"] },
  { key: "operatingIncome", label: "영업이익", aliases: ["영업이익", "영업손실"] },
  { key: "netIncome", label: "당기순이익", aliases: ["당기순이익", "당기순손실", "분기순이익", "분기순손실"] },
];

const state = {
  documents: [],
  activeIndex: 0,
  ocrBusy: false,
};

const el = {
  dropzone: document.querySelector("#dropzone"),
  pdfInput: document.querySelector("#pdfInput"),
  fileTabs: document.querySelector("#fileTabs"),
  resultBody: document.querySelector("#resultBody"),
  rawText: document.querySelector("#rawText"),
  statusText: document.querySelector("#statusText"),
  downloadBtn: document.querySelector("#downloadBtn"),
  ocrBtn: document.querySelector("#ocrBtn"),
  clearBtn: document.querySelector("#clearBtn"),
};

render();

for (const eventName of ["dragenter", "dragover", "dragleave", "drop"]) {
  window.addEventListener(eventName, (event) => {
    event.preventDefault();
    event.stopPropagation();
  });
}

el.dropzone.addEventListener("click", () => el.pdfInput.click());
el.dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") el.pdfInput.click();
});
el.pdfInput.addEventListener("change", (event) => {
  const files = Array.from(event.target.files || []);
  if (files.length) handleFiles(files);
});

for (const eventName of ["dragenter", "dragover"]) {
  el.dropzone.addEventListener(eventName, () => el.dropzone.classList.add("isDragging"));
}

for (const eventName of ["dragleave", "drop"]) {
  el.dropzone.addEventListener(eventName, () => el.dropzone.classList.remove("isDragging"));
}

el.dropzone.addEventListener("drop", (event) => {
  const files = Array.from(event.dataTransfer.files || []).filter((file) => file.name.toLowerCase().endsWith(".pdf"));
  if (files.length) handleFiles(files);
});

el.clearBtn.addEventListener("click", () => {
  state.documents = [];
  state.activeIndex = 0;
  el.pdfInput.value = "";
  el.statusText.textContent = "아직 PDF가 선택되지 않았습니다.";
  render();
});

el.downloadBtn.addEventListener("click", downloadExcel);
el.ocrBtn.addEventListener("click", runOcrForAllDocuments);

async function handleFiles(files) {
  state.documents = files.map(createDocument);
  state.activeIndex = 0;
  render();

  for (let index = 0; index < state.documents.length; index += 1) {
    const doc = state.documents[index];
    el.statusText.textContent = `${doc.file.name} 읽는 중입니다. (${index + 1}/${state.documents.length})`;
    try {
      const text = await extractPdfText(doc.file);
      doc.rawText = text.trim();
      doc.rows = extractFinancialRows(text);
      doc.status = doc.rows.some((row) => row.current || row.previous) ? "텍스트 추출 완료" : "금액 미발견";
    } catch (error) {
      console.error(error);
      doc.status = "PDF 읽기 실패";
    }
    render();
  }

  const foundDocs = state.documents.filter((doc) => doc.rows.some((row) => row.current || row.previous)).length;
  el.statusText.textContent = `${state.documents.length}개 PDF 처리 완료. ${foundDocs}개 파일에서 금액을 찾았습니다.`;
  render();
}

function createDocument(file) {
  return {
    file,
    name: file.name.replace(/\.pdf$/i, ""),
    rows: blankRows(),
    rawText: "",
    status: "대기",
    ocrPdf: null,
  };
}

function blankRows() {
  return TARGET_ITEMS.map((item) => ({ ...item, current: "", previous: "", status: "missing" }));
}

async function extractPdfText(file) {
  const pdfjsLib = await import("./vendor/pdf.min.mjs");
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdf.worker.min.mjs", import.meta.url).toString();

  const data = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const pages = [];

  for (let pageNo = 1; pageNo <= pdf.numPages; pageNo += 1) {
    const page = await pdf.getPage(pageNo);
    const content = await page.getTextContent();
    const items = content.items
      .map((item) => ({ text: item.str, x: item.transform[4], y: Math.round(item.transform[5]) }))
      .filter((item) => item.text.trim());

    items.sort((a, b) => b.y - a.y || a.x - b.x);
    pages.push(`--- ${pageNo}페이지 ---\n${groupIntoLines(items).join("\n")}`);
  }

  return pages.join("\n\n");
}

function groupIntoLines(items) {
  const lines = [];
  for (const item of items) {
    const last = lines[lines.length - 1];
    if (!last || Math.abs(last.y - item.y) > 4) lines.push({ y: item.y, items: [item] });
    else last.items.push(item);
  }

  return lines.map((line) => {
    line.items.sort((a, b) => a.x - b.x);
    return line.items.map((item) => item.text).join(" ").replace(/\s+/g, " ").trim();
  });
}

function extractFinancialRows(text) {
  const lines = text
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);

  return TARGET_ITEMS.map((target) => {
    const match = findBestLine(lines, target);
    if (!match) return { ...target, current: "", previous: "", status: "missing" };

    const isLoss = /손실/.test(match);
    const numbers = extractNumbers(match).map((value) => (isLoss && value > 0 ? -value : value));
    const picked = pickCurrentAndPrevious(numbers);

    return {
      ...target,
      current: picked.current ? formatNumber(picked.current) : "",
      previous: picked.previous ? formatNumber(picked.previous) : "",
      status: picked.current || picked.previous ? "ok" : "warn",
    };
  });
}

function findBestLine(lines, target) {
  const compactAliases = target.aliases.map(normalizeKorean);
  const exact = lines.find((line) => compactAliases.some((alias) => normalizeKorean(line).includes(alias)));
  if (exact) return exact;

  if (target.key === "capital") {
    return lines.find((line) => normalizeKorean(line).includes("자본금") && !normalizeKorean(line).includes("자본금계"));
  }

  return "";
}

function normalizeKorean(value) {
  return value.replace(/\s+/g, "").replace(/[()]/g, "");
}

function extractNumbers(line) {
  return Array.from(line.matchAll(/-?\d[\d,]*\.?\d*/g))
    .map((match) => match[0])
    .filter((raw) => raw.includes(",") || raw.replace(/-/g, "").length >= 4)
    .map(parseMoney)
    .filter((value) => Number.isFinite(value));
}

function pickCurrentAndPrevious(numbers) {
  if (numbers.length >= 4) return { current: numbers[1], previous: numbers[3] };
  if (numbers.length >= 2) return { current: numbers[0], previous: numbers[numbers.length - 1] };
  return { current: numbers[0] || "", previous: "" };
}

function parseMoney(value) {
  if (value === "" || value === null || value === undefined) return "";
  const stringValue = String(value);
  const sign = stringValue.trim().startsWith("-") ? -1 : 1;
  const cleaned = stringValue.replace(/[^0-9.]/g, "");
  if (!cleaned) return "";
  return sign * Number(cleaned);
}

function formatNumber(value) {
  if (value === "" || value === null || Number.isNaN(Number(value))) return "";
  return Number(value).toLocaleString("ko-KR");
}

function render() {
  renderTabs();
  renderRows();
  const hasDocuments = state.documents.length > 0;
  el.downloadBtn.disabled = !hasDocuments || !hasAnyAmount();
  el.ocrBtn.disabled = !hasDocuments || state.ocrBusy;
}

function renderTabs() {
  el.fileTabs.innerHTML = "";
  state.documents.forEach((doc, index) => {
    const button = document.createElement("button");
    button.className = `fileTab${index === state.activeIndex ? " active" : ""}`;
    button.type = "button";
    button.textContent = `${index + 1}. ${doc.file.name}`;
    button.title = doc.status;
    button.addEventListener("click", () => {
      state.activeIndex = index;
      render();
    });
    el.fileTabs.appendChild(button);
  });
}

function renderRows() {
  const doc = activeDocument();
  el.rawText.value = doc?.rawText || "";
  el.resultBody.innerHTML = "";

  const rows = doc?.rows || blankRows();
  rows.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.label}</td>
      <td><input class="moneyInput" data-index="${index}" data-field="current" value="${row.current}" inputmode="numeric" /></td>
      <td><input class="moneyInput" data-index="${index}" data-field="previous" value="${row.previous}" inputmode="numeric" /></td>
      <td>${statusBadge(row)}</td>
    `;
    el.resultBody.appendChild(tr);
  });

  document.querySelectorAll(".moneyInput").forEach((input) => {
    input.disabled = !doc;
    input.addEventListener("input", (event) => {
      const currentDoc = activeDocument();
      if (!currentDoc) return;
      const index = Number(event.target.dataset.index);
      const field = event.target.dataset.field;
      currentDoc.rows[index][field] = event.target.value;
      currentDoc.rows[index].status = currentDoc.rows[index].current || currentDoc.rows[index].previous ? "ok" : "missing";
      el.downloadBtn.disabled = !hasAnyAmount();
    });

    input.addEventListener("blur", (event) => {
      const currentDoc = activeDocument();
      if (!currentDoc) return;
      event.target.value = formatNumber(parseMoney(event.target.value));
      const index = Number(event.target.dataset.index);
      const field = event.target.dataset.field;
      currentDoc.rows[index][field] = event.target.value;
    });
  });
}

function activeDocument() {
  return state.documents[state.activeIndex];
}

function statusBadge(row) {
  if (row.status === "ok") return `<span class="badge ok">확인 필요</span>`;
  if (row.status === "warn") return `<span class="badge warn">금액 없음</span>`;
  return `<span class="badge missing">미입력</span>`;
}

function hasAnyAmount() {
  return state.documents.some((doc) => doc.rows.some((row) => row.current || row.previous));
}

async function runOcrForAllDocuments() {
  if (!window.Tesseract || !window.PDFLib) {
    el.statusText.textContent = "OCR 구성 파일을 불러오지 못했습니다.";
    return;
  }

  state.ocrBusy = true;
  render();

  try {
    for (let index = 0; index < state.documents.length; index += 1) {
      const doc = state.documents[index];
      state.activeIndex = index;
      el.statusText.textContent = `${doc.file.name} OCR 처리 중입니다. 시간이 걸릴 수 있습니다.`;
      render();

      const result = await createOcrPdfForDocument(doc.file, doc.name);
      doc.ocrPdf = result.pdfBytes;
      doc.rawText = result.text.trim() || doc.rawText;
      doc.rows = extractFinancialRows(doc.rawText);
      downloadBlob(
        new Blob([doc.ocrPdf], { type: "application/pdf" }),
        `${doc.name}_OCR완료.pdf`,
      );
      render();
    }
    el.statusText.textContent = "OCR PDF 저장이 완료되었습니다. 추출 결과를 확인한 뒤 엑셀 저장을 누르세요.";
  } catch (error) {
    console.error(error);
    el.statusText.textContent = "OCR 처리 중 문제가 생겼습니다. 인터넷 연결 또는 한국어 OCR 데이터 다운로드 여부를 확인해주세요.";
  } finally {
    state.ocrBusy = false;
    render();
  }
}

async function createOcrPdfForDocument(file, title) {
  const pdfjsLib = await import("./vendor/pdf.min.mjs");
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdf.worker.min.mjs", import.meta.url).toString();

  const worker = await window.Tesseract.createWorker("kor+eng", 1, {
    workerPath: "./vendor/tesseract.worker.min.js",
    logger: (message) => {
      if (message.status && message.progress !== undefined) {
        el.statusText.textContent = `${title} OCR: ${message.status} ${Math.round(message.progress * 100)}%`;
      }
    },
  });
  await worker.setParameters({
    preserve_interword_spaces: "1",
    tessedit_pageseg_mode: "6",
    user_defined_dpi: "300",
  });

  const mergedPdf = await window.PDFLib.PDFDocument.create();
  const original = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  const textParts = [];

  try {
    for (let pageNo = 1; pageNo <= original.numPages; pageNo += 1) {
      const canvas = await renderPdfPageToCanvas(original, pageNo);
      const recognized = await worker.recognize(canvas, { pdfTitle: title }, { pdf: true });
      textParts.push(`--- ${pageNo}페이지 OCR ---\n${normalizeOcrText(recognized.data.text || "")}`);

      if (recognized.data.pdf) {
        const pagePdf = await window.PDFLib.PDFDocument.load(recognized.data.pdf);
        const copied = await mergedPdf.copyPages(pagePdf, pagePdf.getPageIndices());
        copied.forEach((page) => mergedPdf.addPage(page));
      }
    }
  } finally {
    await worker.terminate();
  }

  const pdfBytes = await mergedPdf.save();
  return { pdfBytes, text: textParts.join("\n\n") };
}

async function renderPdfPageToCanvas(pdf, pageNo) {
  const page = await pdf.getPage(pageNo);
  const viewport = page.getViewport({ scale: 3.2 });
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const context = canvas.getContext("2d", { willReadFrequently: true });
  context.fillStyle = "#fff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: context, viewport }).promise;
  return preprocessCanvasForOcr(canvas);
}

function preprocessCanvasForOcr(sourceCanvas) {
  const cropped = cropWhitespace(sourceCanvas);
  const canvas = document.createElement("canvas");
  canvas.width = cropped.width;
  canvas.height = cropped.height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  context.drawImage(cropped, 0, 0);

  const image = context.getImageData(0, 0, canvas.width, canvas.height);
  const data = image.data;
  const gray = new Uint8Array(data.length / 4);

  for (let i = 0, p = 0; i < data.length; i += 4, p += 1) {
    gray[p] = Math.round(data[i] * 0.299 + data[i + 1] * 0.587 + data[i + 2] * 0.114);
  }

  const threshold = otsuThreshold(gray);
  for (let i = 0, p = 0; i < data.length; i += 4, p += 1) {
    const contrast = gray[p] < threshold + 12 ? 0 : 255;
    data[i] = contrast;
    data[i + 1] = contrast;
    data[i + 2] = contrast;
    data[i + 3] = 255;
  }

  context.putImageData(image, 0, 0);
  return canvas;
}

function cropWhitespace(sourceCanvas) {
  const context = sourceCanvas.getContext("2d", { willReadFrequently: true });
  const image = context.getImageData(0, 0, sourceCanvas.width, sourceCanvas.height);
  const data = image.data;
  let minX = sourceCanvas.width;
  let minY = sourceCanvas.height;
  let maxX = 0;
  let maxY = 0;

  for (let y = 0; y < sourceCanvas.height; y += 1) {
    for (let x = 0; x < sourceCanvas.width; x += 1) {
      const index = (y * sourceCanvas.width + x) * 4;
      const gray = data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114;
      if (gray < 245) {
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
      }
    }
  }

  if (maxX <= minX || maxY <= minY) return sourceCanvas;

  const padding = 28;
  minX = Math.max(0, minX - padding);
  minY = Math.max(0, minY - padding);
  maxX = Math.min(sourceCanvas.width, maxX + padding);
  maxY = Math.min(sourceCanvas.height, maxY + padding);

  const canvas = document.createElement("canvas");
  canvas.width = maxX - minX;
  canvas.height = maxY - minY;
  canvas.getContext("2d").drawImage(sourceCanvas, minX, minY, canvas.width, canvas.height, 0, 0, canvas.width, canvas.height);
  return canvas;
}

function otsuThreshold(gray) {
  const histogram = new Array(256).fill(0);
  for (const value of gray) histogram[value] += 1;

  const total = gray.length;
  let sum = 0;
  for (let i = 0; i < 256; i += 1) sum += i * histogram[i];

  let sumB = 0;
  let weightB = 0;
  let maxVariance = 0;
  let threshold = 180;

  for (let i = 0; i < 256; i += 1) {
    weightB += histogram[i];
    if (weightB === 0) continue;
    const weightF = total - weightB;
    if (weightF === 0) break;

    sumB += i * histogram[i];
    const meanB = sumB / weightB;
    const meanF = (sum - sumB) / weightF;
    const variance = weightB * weightF * (meanB - meanF) ** 2;
    if (variance > maxVariance) {
      maxVariance = variance;
      threshold = i;
    }
  }

  return threshold;
}

function normalizeOcrText(text) {
  return text
    .replace(/[０-９]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xfee0))
    .replace(/[，]/g, ",")
    .replace(/[－–—]/g, "-")
    .replace(/영 업/g, "영업")
    .replace(/당 기/g, "당기")
    .replace(/순 이익/g, "순이익")
    .replace(/순 손실/g, "순손실")
    .replace(/자 산/g, "자산")
    .replace(/부 채/g, "부채")
    .replace(/자 본/g, "자본")
    .replace(/[ \t]+/g, " ");
}

function downloadExcel() {
  const sheets = state.documents.map((doc, index) => ({
    name: uniqueSheetName(doc.name || `PDF${index + 1}`, index),
    rows: [
      ["항목", "당기 금액", "전기 금액"],
      ...doc.rows.map((row) => [row.label, parseMoney(row.current) || "", parseMoney(row.previous) || ""]),
    ],
  }));

  const blob = buildXlsxBlob(sheets);
  downloadBlob(blob, `재무제표_추출_${dateStamp()}.xlsx`);
}

function downloadBlob(blob, fileName) {
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function buildXlsxBlob(sheets) {
  const files = [
    {
      name: "[Content_Types].xml",
      text: buildContentTypesXml(sheets.length),
    },
    {
      name: "_rels/.rels",
      text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    },
    {
      name: "xl/workbook.xml",
      text: buildWorkbookXml(sheets),
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      text: buildWorkbookRelsXml(sheets.length),
    },
    {
      name: "xl/styles.xml",
      text: buildStylesXml(),
    },
    ...sheets.map((sheet, index) => ({
      name: `xl/worksheets/sheet${index + 1}.xml`,
      text: buildWorksheetXml(sheet.rows),
    })),
  ];

  return new Blob([zipStore(files)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function buildContentTypesXml(sheetCount) {
  const sheetOverrides = Array.from({ length: sheetCount }, (_, index) =>
    `  <Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
  ).join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${sheetOverrides}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
}

function buildWorkbookXml(sheets) {
  const sheetXml = sheets
    .map((sheet, index) => `    <sheet name="${escapeXmlAttribute(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
${sheetXml}
  </sheets>
</workbook>`;
}

function buildWorkbookRelsXml(sheetCount) {
  const sheetRels = Array.from({ length: sheetCount }, (_, index) =>
    `  <Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
  ).join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheetRels}
  <Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0;[Red]-#,##0"/></numFmts>
  <fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/><color rgb="FFFFFFFF"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF116D6E"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="2"><border/><border><left style="thin"><color rgb="FFD7DCE3"/></left><right style="thin"><color rgb="FFD7DCE3"/></right><top style="thin"><color rgb="FFD7DCE3"/></top><bottom style="thin"><color rgb="FFD7DCE3"/></bottom></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
}

function buildWorksheetXml(rows) {
  const sheetData = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((cell, colIndex) => {
          const ref = `${columnName(colIndex + 1)}${rowIndex + 1}`;
          if (typeof cell === "number") return `<c r="${ref}" s="2"><v>${cell}</v></c>`;
          const style = rowIndex === 0 ? ' s="1"' : "";
          return `<c r="${ref}" t="inlineStr"${style}><is><t>${escapeHtml(cell)}</t></is></c>`;
        })
        .join("");
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="3" width="18" customWidth="1"/></cols>
  <sheetData>${sheetData}</sheetData>
</worksheet>`;
}

function uniqueSheetName(name, index) {
  const clean = name.replace(/[\\/*?:[\]]/g, " ").trim().slice(0, 25) || `PDF${index + 1}`;
  return `${index + 1}_${clean}`.slice(0, 31);
}

function dateStamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;
}

function columnName(index) {
  let name = "";
  while (index > 0) {
    const remainder = (index - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    index = Math.floor((index - 1) / 26);
  }
  return name;
}

function zipStore(files) {
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const file of files) {
    const nameBytes = encoder.encode(file.name);
    const data = encoder.encode(file.text);
    const crc = crc32(data);

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(10, dosTime(), true);
    localView.setUint16(12, dosDate(), true);
    localView.setUint32(14, crc, true);
    localView.setUint32(18, data.length, true);
    localView.setUint32(22, data.length, true);
    localView.setUint16(26, nameBytes.length, true);
    localHeader.set(nameBytes, 30);
    localParts.push(localHeader, data);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(12, dosTime(), true);
    centralView.setUint16(14, dosDate(), true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(20, data.length, true);
    centralView.setUint32(24, data.length, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);

    offset += localHeader.length + data.length;
  }

  const centralOffset = offset;
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, files.length, true);
  endView.setUint16(10, files.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, centralOffset, true);

  return concatUint8Arrays([...localParts, ...centralParts, end]);
}

function concatUint8Arrays(parts) {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const output = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function dosTime() {
  const now = new Date();
  return (now.getHours() << 11) | (now.getMinutes() << 5) | Math.floor(now.getSeconds() / 2);
}

function dosDate() {
  const now = new Date();
  return ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
}

function crc32(data) {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ byte) & 0xff];
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeXmlAttribute(value) {
  return escapeHtml(value).replace(/'/g, "&apos;");
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let value = i;
    for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    table[i] = value >>> 0;
  }
  return table;
})();
