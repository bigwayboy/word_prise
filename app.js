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
const resultTitle = document.querySelector("#result-title");
const resultCard = document.querySelector("#result-card");
const analysisCard = document.querySelector("#analysis-card");
const metricGrid = document.querySelector("#metric-grid");
const rushInput = document.querySelector("#rush");
const dropZone = document.querySelector("#drop-zone");
const quoteDetailButton = document.querySelector("#quote-detail-button");
const detailList = document.querySelector("#detail-list");

let latestDetails = {};
let latestQuoteRows = [];
let latestRush = false;

renderOptions();

uploadButton.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  setSelectedFile(file);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("is-dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    if (event.target === dropZone || eventName === "drop") {
      dropZone.classList.remove("is-dragging");
    }
  });
});

dropZone.addEventListener("drop", (event) => {
  const file = Array.from(event.dataTransfer?.files || []).find((item) =>
    /\.(doc|docx)$/i.test(item.name),
  );

  if (!file) {
    setStatus("请拖入 .doc 或 .docx 格式的 Word 文件。", true);
    return;
  }

  const dataTransfer = new DataTransfer();
  dataTransfer.items.add(file);
  fileInput.files = dataTransfer.files;
  setSelectedFile(file);
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

quoteDetailButton.addEventListener("click", () => {
  renderQuote(latestQuoteRows, latestRush);
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
    latestDetails = metrics.details || {};
    latestQuoteRows = rows;
    latestRush = rushInput.checked;
    renderMetrics(metrics);
    renderQuote(latestQuoteRows, latestRush);
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

function setSelectedFile(file) {
  fileName.textContent = file ? file.name : "尚未选择文件";
  clearStatus();
  hideResults();
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
    ["字数", metrics.wordCount, null],
    ["图片", metrics.imageCount, "images"],
    ["表格", metrics.tableCount, "tables"],
    ["参考文献", metrics.referenceCount, "references"],
    ["公式", metrics.formulaCount, "formulas"],
    ["脚注", metrics.footnoteCount, "footnotes"],
  ];

  metricGrid.innerHTML = items
    .map(
      ([label, value, detailKey]) => `
        <button class="metric ${detailKey ? "is-clickable" : ""}" type="button" ${
          detailKey ? `data-detail-key="${detailKey}"` : "disabled"
        }>
          <strong>${value}</strong>
          <span>${label}</span>
        </button>
      `,
    )
    .join("");

  metricGrid.querySelectorAll("[data-detail-key]").forEach((button) => {
    button.addEventListener("click", () => {
      renderDetailList(button.dataset.detailKey, button.querySelector("span").textContent);
    });
  });
  analysisCard.classList.remove("is-hidden");
}

function renderQuote(rows, isRush) {
  resultTitle.textContent = "报价明细";
  quoteDetailButton.classList.add("is-hidden");
  detailList.classList.add("is-hidden");
  detailList.innerHTML = "";
  document.querySelector(".table-wrap").classList.remove("is-hidden");
  totalPrice.classList.remove("is-hidden");
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

function renderDetailList(detailKey, label) {
  const details = latestDetails?.[detailKey] || [];
  resultTitle.textContent = `${label}明细`;
  quoteDetailButton.classList.remove("is-hidden");
  document.querySelector(".table-wrap").classList.add("is-hidden");
  totalPrice.classList.add("is-hidden");
  detailList.classList.remove("is-hidden");
  detailList.innerHTML =
    details.length > 0
      ? `<ol>${details.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ol>`
      : `<p class="empty-detail">没有检测到${label}。</p>`;
  resultCard.classList.remove("is-hidden");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
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
  detailList.innerHTML = "";
  detailList.classList.add("is-hidden");
  quoteDetailButton.classList.add("is-hidden");
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
  const appXml = zipEntries.get("docProps/app.xml") || "";
  const paragraphs = extractParagraphs(documentXml);
  const wordCount = countDocumentWords(zipEntries, appXml);
  const imageDetails = getImageDetails(documentXml, zipEntries);
  const tableDetails = getTableDetails(documentXml);
  const footnoteDetails = getFootnoteDetails(footnotesXml);
  const referenceDetails = getReferenceTypeMarkerDetails(paragraphs);
  const formulaDetails = getFormulaDetails(documentXml);

  return {
    wordCount,
    imageCount: imageDetails.length,
    tableCount: tableDetails.length,
    figureTableCount: imageDetails.length + tableDetails.length,
    footnoteCount: footnoteDetails.length,
    referenceCount: referenceDetails.length,
    formulaCount: formulaDetails.length,
    details: {
      images: imageDetails,
      tables: tableDetails,
      references: referenceDetails,
      formulas: formulaDetails,
      footnotes: footnoteDetails,
    },
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

    if (name.endsWith(".xml") || name.endsWith(".rels") || name.startsWith("word/media/")) {
      const bytes = await readZipEntryBytes(
        buffer,
        view,
        localHeaderOffset,
        compressedSize,
        compressionMethod,
      );
      entries.set(name, name.endsWith(".xml") || name.endsWith(".rels") ? textDecoder.decode(bytes) : bytes);
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

function countDocumentWords(zipEntries, appXml) {
  const calculatedWords = countWpsLikeWords(extractCountableDocumentText(zipEntries));
  const savedWords = getSavedWordCount(appXml);
  return savedWords === null ? calculatedWords : Math.max(calculatedWords, savedWords);
}

function extractCountableDocumentText(zipEntries) {
  return Array.from(zipEntries.entries())
    .filter(([name, value]) => typeof value === "string" && isCountableWordXml(name))
    .map(([, xml]) => xmlTextForWordCount(xml))
    .join("");
}

function isCountableWordXml(name) {
  return /^word\/(?:document|footnotes|endnotes|header\d+|footer\d+)\.xml$/i.test(name);
}

function xmlTextForWordCount(xml) {
  return Array.from(xml.matchAll(/<(?:w|a|m):t(?:\s[^>]*)?>([\s\S]*?)<\/(?:w|a|m):t>/g))
    .map((match) => decodeXml(match[1]))
    .join("");
}

function getSavedWordCount(appXml) {
  const savedWords = getXmlNumber(appXml, "Words");
  if (savedWords !== null) {
    return savedWords;
  }

  return getXmlNumber(appXml, "Characters");
}

function getXmlNumber(xml, tagName) {
  const match = xml.match(new RegExp(`<${tagName}>(\\d+)<\\/${tagName}>`, "i"));
  if (!match) {
    return null;
  }

  const value = Number(match[1]);
  return Number.isFinite(value) && value >= 0 ? value : null;
}

function countWpsLikeWords(text) {
  const normalized = text.normalize("NFKC");
  const chineseChars = normalized.match(/[\u3400-\u9fff]/g)?.length || 0;
  const latinWords = normalized.match(/[A-Za-z]+(?:[-'][A-Za-z]+)*/g)?.length || 0;
  const numberTokens = normalized.match(/\d+(?:[.,]\d+)*/g)?.length || 0;
  const digitChars = normalized.match(/\d/g)?.length || 0;
  const selectedPunctuation = normalized.match(/[,.，。:：;；()（）\-—%％]/g)?.length || 0;
  const baseCount = chineseChars + latinWords + numberTokens;
  const digitExtra = Math.max(0, digitChars - numberTokens);
  let total = baseCount;

  if (baseCount > 0 && digitExtra / baseCount <= 0.08) {
    total += digitExtra;
  }

  if (baseCount > 0 && selectedPunctuation / baseCount <= 0.05) {
    total += selectedPunctuation;
  }

  return total;
}

function countImages(documentXml, zipEntries) {
  return getImageDetails(documentXml, zipEntries).length;
}

function getImageDetails(documentXml, zipEntries) {
  const relationshipMap = parseRelationshipTargets(zipEntries.get("word/_rels/document.xml.rels") || "");
  return getRealImageDetails(documentXml, relationshipMap);
}

function parseRelationshipTargets(relsXml) {
  const targets = new Map();
  Array.from(relsXml.matchAll(/<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/g)).forEach(
    (match) => {
      targets.set(match[1], match[2]);
    },
  );
  return targets;
}

function countRealImageReferences(documentXml, relationshipMap) {
  return getRealImageDetails(documentXml, relationshipMap).length;
}

function getRealImageDetails(documentXml, relationshipMap) {
  const paragraphXmlList = Array.from(documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)).map((match) => match[0]);
  const details = [];

  paragraphXmlList.forEach((paragraphXml, paragraphIndex) => {
    const imageRefs = extractParagraphImageRefs(paragraphXml, relationshipMap);
    imageRefs.forEach((imageRef) => {
      if (isRealFigureImage(imageRef)) {
        details.push(
          `图片${details.length + 1}：${imageRef.extension.toUpperCase()}，约在第 ${paragraphIndex + 1} 段`,
        );
      }
    });
  });

  return details;
}

function extractParagraphImageRefs(paragraphXml, relationshipMap) {
  const refs = [];
  const drawingBlocks = paragraphXml.match(/<w:drawing[\s\S]*?<\/w:drawing>/g) || [];
  const pictBlocks = paragraphXml.match(/<w:pict[\s\S]*?<\/w:pict>/g) || [];

  drawingBlocks.forEach((block) => {
    const extension = getImageExtension(block, relationshipMap);
    const extent = block.match(/<wp:extent[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
    refs.push({
      extension,
      width: extent ? Number(extent[1]) : 0,
      height: extent ? Number(extent[2]) : 0,
    });
  });

  pictBlocks.forEach((block) => {
    refs.push({
      extension: getImageExtension(block, relationshipMap),
      width: 0,
      height: 0,
    });
  });

  return refs;
}

function getImageExtension(xml, relationshipMap) {
  const relationshipIdMatch = xml.match(/r:(?:embed|id)="([^"]+)"/);
  const target = relationshipIdMatch ? relationshipMap.get(relationshipIdMatch[1]) || "" : "";
  const mediaMatch = target.match(/media\/[^"']+\.([A-Za-z0-9]+)/i);
  return mediaMatch ? mediaMatch[1].toLowerCase() : "";
}

function isRealFigureImage(imageRef) {
  if (!imageRef.extension) {
    return false;
  }

  if (!["wmf", "emf"].includes(imageRef.extension)) {
    return true;
  }

  return imageRef.height >= 900000;
}

function countFootnotes(xml) {
  return getFootnoteDetails(xml).length;
}

function getFootnoteDetails(xml) {
  if (!xml) {
    return [];
  }

  return Array.from(xml.matchAll(/<w:footnote\b([^>]*)>([\s\S]*?)<\/w:footnote>/g))
    .filter((match) => !/w:type="(?:separator|continuationSeparator|continuationNotice)"/.test(match[1]))
    .map((match, index) => {
      const text = normalizeText(xmlParagraphToText(match[2])) || "未提取到脚注文字";
      return `脚注${index + 1}：${truncateText(text, 80)}`;
    });
}

function getTableDetails(documentXml) {
  return Array.from(documentXml.matchAll(/<w:tbl[\s\S]*?<\/w:tbl>/g)).map((match, index) => {
    const text = normalizeText(xmlParagraphToText(match[0])) || "未提取到表格文字";
    return `表格${index + 1}：${truncateText(text, 80)}`;
  });
}

function getReferenceSectionText(paragraphs) {
  const referenceHeadingIndex = paragraphs.findLastIndex((text) =>
    /^(参考文献|参考资料|References|Bibliography)\b[:：]?$|^参考文献[:：]/i.test(text.trim()),
  );
  const referenceParagraphs =
    referenceHeadingIndex >= 0 ? paragraphs.slice(referenceHeadingIndex + 1) : paragraphs;
  return referenceParagraphs.join(" ");
}

function countReferenceTypeMarkers(paragraphs) {
  return getReferenceTypeMarkerDetails(paragraphs).length;
}

function getReferenceTypeMarkerDetails(paragraphs) {
  const referenceParagraphs = getReferenceSectionParagraphs(paragraphs);
  const markerPattern =
    /[［\[](?:M\/OL|J\/OL|C\/OL|R\/OL|D\/OL|P\/OL|S\/OL|N\/OL|EB\/OL|J|D|C|S|M|G|R|P|N|EB|A|Z)[］\]]/gi;
  const details = [];

  referenceParagraphs.forEach((paragraph) => {
    const normalized = paragraph.replace(/\s+/g, "");
    const markers = normalized.match(markerPattern) || [];
    markers.forEach((marker) => {
      details.push(`参考文献${details.length + 1}：${marker} ${truncateText(paragraph, 90)}`);
    });
  });

  return details;
}

function getReferenceSectionParagraphs(paragraphs) {
  const referenceHeadingIndex = paragraphs.findLastIndex((text) =>
    /^(参考文献|参考资料|References|Bibliography)\b[:：]?$|^参考文献[:：]/i.test(text.trim()),
  );
  return referenceHeadingIndex >= 0 ? paragraphs.slice(referenceHeadingIndex + 1) : paragraphs;
}

function countFormulas(documentXml, paragraphs) {
  return getFormulaDetails(documentXml).length;
}

function getFormulaDetails(documentXml) {
  const paragraphXmlList = Array.from(documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)).map((match) => match[0]);
  const details = [];

  paragraphXmlList.forEach((paragraphXml, paragraphIndex) => {
    const text = normalizeText(xmlParagraphToText(paragraphXml));
    const markerCount = countEquationNumberMarkers(text);

    if (markerCount > 0) {
      for (let index = 0; index < markerCount; index += 1) {
        details.push(`公式${details.length + 1}：检测到公式编号，约在第 ${paragraphIndex + 1} 段`);
      }
      return;
    }

    const mathTypeCount = countMathTypeObjects(paragraphXml);
    if (mathTypeCount > 0) {
      for (let index = 0; index < mathTypeCount; index += 1) {
        details.push(`公式${details.length + 1}：MathType/Equation 对象，约在第 ${paragraphIndex + 1} 段`);
      }
      return;
    }

    if (hasStandaloneOfficeFormula(paragraphXml, text)) {
      details.push(`公式${details.length + 1}：Office 公式，约在第 ${paragraphIndex + 1} 段`);
    }
  });

  return details;
}

function countEquationNumberMarkers(text) {
  const normalized = text
    .normalize("NFKC")
    .replace(/[－—–]/g, "-")
    .replace(/[．。]/g, ".");
  const matches = normalized.match(/\(\s*(?:式\s*)?\d+\s*(?:[.-]\s*\d+)+\s*\)/g);
  return matches?.length || 0;
}

function countMathTypeObjects(paragraphXml) {
  return paragraphXml.match(/<o:OLEObject\b[^>]*\bProgID="(?:Equation|MathType)[^"]*"/gi)?.length || 0;
}

function hasStandaloneOfficeFormula(paragraphXml, text) {
  const hasOfficeMath = /<m:oMath|<m:oMathPara/.test(paragraphXml);
  const hasFormulaText = /[=≈≠≤≥∑∫√∞±×÷α-ωΑ-Ω]|\\frac|\\sum|\\int|lim|sin|cos|tan|log/i.test(text);
  const isStandalone = text.length <= 160;
  return isStandalone && hasOfficeMath && hasFormulaText;
}

function countMatches(text, regex) {
  return text.match(regex)?.length || 0;
}

function truncateText(text, maxLength) {
  return text.length > maxLength ? `${text.slice(0, maxLength)}...` : text;
}
