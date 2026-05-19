const PRICE_CONFIG = {
  textFormatPerThousand: 4,
  toc: 15,
  cover: 10,
  headerFooter: 15,
  figuresTables: 0.9,
  references: 1.2,
  formulas: 0.9,
  footnotes: 0.8,
  rushRate: 0.3,
  freeCoverThreshold: 100,
};

const PRICING_ITEMS = [
  {
    id: "textFormat",
    label: "标题及正文格式",
    unitLabel: `${PRICE_CONFIG.textFormatPerThousand} 元 / 千字`,
    price: PRICE_CONFIG.textFormatPerThousand,
    quantity: (metrics) => metrics.wordCount / 1000,
    quantityLabel: (metrics) => `${metrics.wordCount} 字`,
  },
  {
    id: "toc",
    label: "目录",
    unitLabel: `${PRICE_CONFIG.toc} 元 / 项`,
    price: PRICE_CONFIG.toc,
    quantity: () => 1,
    quantityLabel: () => "1 项",
  },
  {
    id: "cover",
    label: "封面",
    unitLabel: `${PRICE_CONFIG.cover} 元 / 项`,
    price: PRICE_CONFIG.cover,
    quantity: () => 1,
    quantityLabel: () => "1 项",
  },
  {
    id: "headerFooter",
    label: "页眉页脚",
    unitLabel: `${PRICE_CONFIG.headerFooter} 元 / 项`,
    price: PRICE_CONFIG.headerFooter,
    quantity: () => 1,
    quantityLabel: () => "1 项",
  },
  {
    id: "figuresTables",
    label: "图表",
    unitLabel: `${PRICE_CONFIG.figuresTables} 元 / 个`,
    price: PRICE_CONFIG.figuresTables,
    quantity: (metrics) => metrics.figureTableCount,
    quantityLabel: (metrics) =>
      `${metrics.figureTableCount} 个（图片 ${metrics.imageCount}，表格 ${metrics.tableCount}）`,
  },
  {
    id: "references",
    label: "参考文献引用格式",
    unitLabel: `${PRICE_CONFIG.references} 元 / 条`,
    price: PRICE_CONFIG.references,
    quantity: (metrics) => metrics.referenceCount,
    quantityLabel: (metrics) => `${metrics.referenceCount} 条`,
  },
  {
    id: "formulas",
    label: "公式（准确性不高）",
    unitLabel: `${PRICE_CONFIG.formulas} 元 / 个`,
    price: PRICE_CONFIG.formulas,
    checkedByDefault: false,
    quantity: (metrics) => metrics.formulaCount,
    quantityLabel: (metrics) => `${metrics.formulaCount} 个`,
  },
  {
    id: "footnotes",
    label: "脚注",
    unitLabel: `${PRICE_CONFIG.footnotes} 元 / 个`,
    price: PRICE_CONFIG.footnotes,
    quantity: (metrics) => metrics.footnoteCount,
    quantityLabel: (metrics) => `${metrics.footnoteCount} 个`,
  },
];

const ZIP_LOCAL_FILE_HEADER = 0x04034b50;
const ZIP_CENTRAL_DIRECTORY_HEADER = 0x02014b50;
const ZIP_END_OF_CENTRAL_DIRECTORY = 0x06054b50;
const textDecoder = new TextDecoder("utf-8");

const fileInput = document.querySelector("#word-file");
const uploadButton = document.querySelector("#upload-button");
const fileName = document.querySelector("#file-name");
const statusMessage = document.querySelector("#status-message");
const priceOptions = document.querySelector("#price-options");
const quoteForm = document.querySelector("#quote-form");
const quoteBody = document.querySelector("#quote-body");
const totalPrice = document.querySelector("#total-price");
const resultCard = document.querySelector("#result-card");
const analysisCard = document.querySelector("#analysis-card");
const metricGrid = document.querySelector("#metric-grid");
const rushInput = document.querySelector("#rush");

renderOptions();

uploadButton.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  fileName.textContent = file ? file.name : "尚未选择文件";
  clearStatus();
  hideResults();
});

document.querySelector("#select-all").addEventListener("click", () => {
  document.querySelectorAll(".item-checkbox").forEach((input) => {
    input.checked = true;
  });
});

document.querySelector("#clear-all").addEventListener("click", () => {
  document.querySelectorAll(".item-checkbox").forEach((input) => {
    input.checked = false;
  });
});

quoteForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  hideResults();
  clearStatus();

  const file = fileInput.files?.[0];
  if (!file) {
    setStatus("请先上传 Word 文件。", true);
    return;
  }

  const lowerName = file.name.toLowerCase();
  if (lowerName.endsWith(".doc") && !lowerName.endsWith(".docx")) {
    setStatus("当前版本无法自动解析 .doc 老格式，请将文件另存为 .docx 后重新上传。", true);
    return;
  }

  if (!lowerName.endsWith(".docx")) {
    setStatus("请上传 .doc 或 .docx 格式的 Word 文件。", true);
    return;
  }

  const selectedIds = getSelectedItemIds();
  if (selectedIds.length === 0) {
    setStatus("请至少选择一个需要修改的项目。", true);
    return;
  }

  setStatus("正在分析文档，请稍候...");

  try {
    const metrics = await analyzeDocx(file);
    const rows = buildQuoteRows(metrics, selectedIds);
    renderMetrics(metrics);
    renderQuote(rows, rushInput.checked);
    setStatus("报价已生成。");
  } catch (error) {
    console.error(error);
    setStatus(`文档解析失败：${error.message || "请确认文件是有效的 .docx 文档。"}`, true);
  }
});

function renderOptions() {
  priceOptions.innerHTML = PRICING_ITEMS.map(
    (item) => `
      <label class="price-option">
        <input class="item-checkbox" type="checkbox" value="${item.id}" ${
          item.checkedByDefault === false ? "" : "checked"
        } />
        <span>
          <strong>${item.label}</strong>
          <small>${item.unitLabel}</small>
        </span>
        <em>${item.unitLabel}</em>
      </label>
    `,
  ).join("");
}

function getSelectedItemIds() {
  return Array.from(document.querySelectorAll(".item-checkbox:checked")).map((input) => input.value);
}

function buildQuoteRows(metrics, selectedIds) {
  return PRICING_ITEMS.filter((item) => selectedIds.includes(item.id)).map((item) => {
    const quantity = item.quantity(metrics);
    return {
      id: item.id,
      label: item.label,
      quantity,
      quantityLabel: item.quantityLabel(metrics),
      unitLabel: item.unitLabel,
      subtotal: quantity * item.price,
    };
  });
}

function renderMetrics(metrics) {
  const items = [
    ["字数", metrics.wordCount],
    ["图片", metrics.imageCount],
    ["表格", metrics.tableCount],
    ["参考文献", metrics.referenceCount],
    ["公式", metrics.formulaCount],
    ["脚注", metrics.footnoteCount],
  ];

  metricGrid.innerHTML = items
    .map(
      ([label, value]) => `
        <div class="metric">
          <strong>${value}</strong>
          <span>${label}</span>
        </div>
      `,
    )
    .join("");
  analysisCard.classList.remove("is-hidden");
}

function renderQuote(rows, isRush) {
  const originalTotal = rows.reduce((sum, row) => sum + row.subtotal, 0);
  const shouldFreeCover =
    originalTotal > PRICE_CONFIG.freeCoverThreshold && rows.some((row) => row.id === "cover");
  const discount = shouldFreeCover ? PRICE_CONFIG.cover : 0;
  const discountedTotal = Math.max(0, originalTotal - discount);
  const rushFee = isRush ? discountedTotal * PRICE_CONFIG.rushRate : 0;
  const displayRows = rows.map((row) => ({
    ...row,
    isFreeCover: shouldFreeCover && row.id === "cover",
  }));

  if (isRush) {
    displayRows.push({
      label: "加急",
      quantityLabel: "总价 × 30%",
      unitLabel: `${PRICE_CONFIG.rushRate * 100}%`,
      subtotal: rushFee,
    });
  }

  quoteBody.innerHTML = displayRows
    .map(
      (row) => `
        <tr>
          <td data-label="项目">${row.label}${row.isFreeCover ? "（满减优惠）" : ""}</td>
          <td data-label="数量">${row.quantityLabel}</td>
          <td data-label="单价">${row.unitLabel}</td>
          <td data-label="小计">${
            row.isFreeCover ? `<s>${formatMoney(row.subtotal)}</s>` : formatMoney(row.subtotal)
          }</td>
        </tr>
      `,
    )
    .join("");

  totalPrice.textContent = formatMoney(discountedTotal + rushFee);
  resultCard.classList.remove("is-hidden");
}

function formatMoney(value) {
  return `¥${value.toFixed(2)}`;
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("is-error", isError);
}

function clearStatus() {
  setStatus("");
}

function hideResults() {
  resultCard.classList.add("is-hidden");
  analysisCard.classList.add("is-hidden");
  quoteBody.innerHTML = "";
  metricGrid.innerHTML = "";
  totalPrice.textContent = "¥0.00";
}

async function analyzeDocx(file) {
  const buffer = await file.arrayBuffer();
  const zipEntries = await readZipEntries(buffer);
  const documentXml = zipEntries.get("word/document.xml");

  if (!documentXml) {
    throw new Error("未找到 word/document.xml，文件可能不是标准 .docx。");
  }

  const footnotesXml = zipEntries.get("word/footnotes.xml") || "";
  const paragraphs = extractParagraphs(documentXml);
  const plainText = paragraphs.join("");
  const bodyText = getBodyTextBeforeReferences(paragraphs);
  const imageCount = countImages(documentXml, zipEntries);
  const tableCount = countMatches(documentXml, /<w:tbl[\s>]/g);
  const footnoteCount = countFootnotes(footnotesXml);
  const referenceCount = countCitationMarkers(bodyText);
  const formulaCount = countFormulas(documentXml, paragraphs);

  return {
    wordCount: countTextCharacters(plainText),
    imageCount,
    tableCount,
    figureTableCount: imageCount + tableCount,
    footnoteCount,
    referenceCount,
    formulaCount,
  };
}

async function readZipEntries(buffer) {
  const view = new DataView(buffer);
  const entries = new Map();
  const eocdOffset = findEndOfCentralDirectory(view);

  if (eocdOffset < 0) {
    throw new Error("无法读取 docx 压缩结构。");
  }

  const centralDirectorySize = view.getUint32(eocdOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  let offset = centralDirectoryOffset;
  const end = centralDirectoryOffset + centralDirectorySize;

  while (offset < end && view.getUint32(offset, true) === ZIP_CENTRAL_DIRECTORY_HEADER) {
    const compressionMethod = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const nameStart = offset + 46;
    const rawName = new Uint8Array(buffer, nameStart, fileNameLength);
    const name = textDecoder.decode(rawName).replaceAll("\\", "/");

    if (name.endsWith(".xml") || name.startsWith("word/media/")) {
      const bytes = await readZipEntryBytes(
        buffer,
        view,
        localHeaderOffset,
        compressedSize,
        compressionMethod,
      );
      entries.set(name, name.endsWith(".xml") ? textDecoder.decode(bytes) : bytes);
    }

    offset = nameStart + fileNameLength + extraLength + commentLength;
  }

  return entries;
}

function findEndOfCentralDirectory(view) {
  const minOffset = Math.max(0, view.byteLength - 65557);
  for (let offset = view.byteLength - 22; offset >= minOffset; offset -= 1) {
    if (view.getUint32(offset, true) === ZIP_END_OF_CENTRAL_DIRECTORY) {
      return offset;
    }
  }
  return -1;
}

async function readZipEntryBytes(buffer, view, localHeaderOffset, compressedSize, compressionMethod) {
  if (view.getUint32(localHeaderOffset, true) !== ZIP_LOCAL_FILE_HEADER) {
    throw new Error("文件压缩结构异常。");
  }

  const fileNameLength = view.getUint16(localHeaderOffset + 26, true);
  const extraLength = view.getUint16(localHeaderOffset + 28, true);
  const dataStart = localHeaderOffset + 30 + fileNameLength + extraLength;
  const compressedBytes = new Uint8Array(buffer, dataStart, compressedSize);

  if (compressionMethod === 0) {
    return compressedBytes;
  }

  if (compressionMethod === 8) {
    return inflateRaw(compressedBytes);
  }

  throw new Error(`暂不支持此 docx 压缩方式：${compressionMethod}`);
}

async function inflateRaw(bytes) {
  if (!("DecompressionStream" in window)) {
    throw new Error("当前浏览器不支持本地解压，请使用最新版 Chrome、Edge 或 Safari。");
  }

  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  const response = new Response(stream);
  return new Uint8Array(await response.arrayBuffer());
}

function extractParagraphs(xml) {
  return Array.from(xml.matchAll(/<w:p[\s\S]*?<\/w:p>/g))
    .map((match) => xmlParagraphToText(match[0]))
    .map((text) => normalizeText(text))
    .filter(Boolean);
}

function xmlParagraphToText(paragraphXml) {
  return Array.from(paragraphXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g))
    .map((match) => decodeXml(match[1]))
    .join("");
}

function decodeXml(value) {
  const textarea = document.createElement("textarea");
  textarea.innerHTML = value;
  return textarea.value;
}

function normalizeText(value) {
  return value.replace(/\s+/g, " ").trim();
}

function countTextCharacters(text) {
  const withoutReferences = text.replace(/\s+/g, "");
  const chineseChars = withoutReferences.match(/[\u3400-\u9fff]/g)?.length || 0;
  const latinWords = withoutReferences.match(/[A-Za-z0-9]+(?:[-'][A-Za-z0-9]+)*/g)?.length || 0;
  return chineseChars + latinWords;
}

function countImages(documentXml, zipEntries) {
  const mediaCount = Array.from(zipEntries.keys()).filter((name) => name.startsWith("word/media/")).length;
  const drawingCount = countMatches(documentXml, /<(?:w:)?drawing[\s>]/g);
  const pictCount = countMatches(documentXml, /<(?:w:)?pict[\s>]/g);
  return Math.max(mediaCount, drawingCount + pictCount);
}

function countFootnotes(xml) {
  if (!xml) {
    return 0;
  }

  return Array.from(xml.matchAll(/<w:footnote\b([^>]*)>/g)).filter((match) => {
    const attrs = match[1];
    return !/w:type="(?:separator|continuationSeparator|continuationNotice)"/.test(attrs);
  }).length;
}

function getBodyTextBeforeReferences(paragraphs) {
  const referenceHeadingIndex = paragraphs.findLastIndex((text) =>
    /^(参考文献|参考资料|References|Bibliography)\b[:：]?$|^参考文献[:：]/i.test(text.trim()),
  );
  const bodyParagraphs =
    referenceHeadingIndex > 0 ? paragraphs.slice(0, referenceHeadingIndex) : paragraphs;
  return bodyParagraphs.join(" ");
}

function countCitationMarkers(text) {
  const citationNumbers = new Set();
  const normalized = text.replace(/\s+/g, "");
  const markerPattern = /[［\[]([0-9,;\-－—–，、；~～\s]+)[］\]]/g;
  let match;

  while ((match = markerPattern.exec(normalized)) !== null) {
    expandCitationNumbers(match[1]).forEach((number) => citationNumbers.add(number));
  }

  return citationNumbers.size;
}

function expandCitationNumbers(markerText) {
  const normalized = markerText
    .replace(/[，、]/g, ",")
    .replace(/[；;]/g, ",")
    .replace(/[－—–~～]/g, "-")
    .replace(/\s+/g, "");
  const numbers = new Set();

  normalized.split(",").forEach((part) => {
    if (!part) {
      return;
    }

    const rangeMatch = part.match(/^(\d+)-(\d+)$/);
    if (rangeMatch) {
      const start = Number(rangeMatch[1]);
      const end = Number(rangeMatch[2]);
      if (Number.isInteger(start) && Number.isInteger(end) && end >= start && end - start <= 200) {
        for (let number = start; number <= end; number += 1) {
          numbers.add(number);
        }
      }
      return;
    }

    const number = Number(part);
    if (Number.isInteger(number) && number > 0) {
      numbers.add(number);
    }
  });

  return numbers;
}

function countFormulas(documentXml, paragraphs) {
  const paragraphXmlList = Array.from(documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)).map((match) => match[0]);
  let count = 0;

  paragraphXmlList.forEach((paragraphXml) => {
    const text = normalizeText(xmlParagraphToText(paragraphXml));
    const hasOfficeMath = /<m:oMath|<m:oMathPara/.test(paragraphXml);
    const hasFormulaText = /[=≈≠≤≥∑∫√∞±×÷α-ωΑ-Ω]|\\frac|\\sum|\\int|lim|sin|cos|tan|log/i.test(text);
    const hasNumber = /(?:\(?（?\d+(?:\.\d+)*\)?）?|式\s*\d+)\s*$/.test(text);
    const isStandalone = text.length <= 160;

    if (isStandalone && hasNumber && (hasOfficeMath || hasFormulaText)) {
      count += 1;
    }
  });

  return count;
}

function countMatches(text, regex) {
  return text.match(regex)?.length || 0;
}
