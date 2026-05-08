const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Header,
  Footer,
  AlignmentType,
  BorderStyle,
  PageBreak,
} = require("docx");
const JSZip = require("jszip");

const DEFAULT_TEMPLATE_PATH = path.join(__dirname, "..", "templates", "岭南学院公文模板.docx");

const FONT_HEADER = "\u65B9\u6B63\u5C0F\u6807\u5B8B\u7B80\u4F53";
const FONT_BODY = "\u4EFF\u5B8B_GB2312";
const FONT_LEVEL1 = "\u9ED1\u4F53";
const FONT_ASCII = "Times New Roman";
const FONT_DISCLOSURE_KEY = "\u9ED1\u4F53";
const FONT_DISCLOSURE_VALUE = "\u4EFF\u5B8B_GB2312";

const SIZE_HEADER = 44;
const SIZE_TITLE = 22;
const SIZE_BODY = 16;
const SIZE_FOOTER = 14;
const SIZE_DISCLOSURE = 14;

const PAGE_WIDTH = 11906;
const PAGE_HEIGHT = 16838;
const MARGIN_TOP = 2098;
const MARGIN_BOTTOM = 2041;
const MARGIN_LEFT = 1588;
const MARGIN_RIGHT = 1588;
const HEADER_DISTANCE = 851;
const FOOTER_DISTANCE = 1644;

const LINE_SPACING = 540;
const TWO_CHAR_INDENT = 640;
const RIGHT_INDENT_SIGNER = 1247;
const RIGHT_INDENT_DATE = 10;

const CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT;

function twipsToHalfPt(twips) {
  return twips;
}

function ptToHalfPt(pt) {
  return pt * 2;
}

function normalizeQuotes(text) {
  if (!text) return text;
  return text.replace(/"([^"\n]+)"/g, "\u201C$1\u201D");
}

function normalizeRecipients(value) {
  if (!value) return "";
  if (Array.isArray(value)) {
    return value.map((s) => String(s).trim()).filter(Boolean).join("\u3001");
  }
  return String(value).trim();
}

function normalizeList(value) {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value.map((s) => String(s).trim()).filter(Boolean);
  }
  const text = String(value).trim();
  return text ? [text] : [];
}

function parseBodyText(text) {
  if (!text) return [];
  return text
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);
}

function parseSimpleFrontMatter(text) {
  const stripped = text.replace(/^\uFEFF/, "");
  if (!stripped.startsWith("---\n") && stripped !== "---") {
    return { meta: {}, rest: text };
  }
  const lines = stripped.split("\n");
  if (!lines.length || lines[0].trim() !== "---") {
    return { meta: {}, rest: text };
  }

  const meta = {};
  let i = 1;
  let currentListKey = null;
  while (i < lines.length) {
    const line = lines[i];
    if (line.trim() === "---") {
      i++;
      break;
    }
    if (!line.trim() || line.trimStart().startsWith("#")) {
      i++;
      continue;
    }
    if (line.startsWith("  - ") || line.startsWith("- ")) {
      if (currentListKey) {
        const item = line.split("-").slice(1).join("-").trim();
        if (item) {
          if (!meta[currentListKey]) meta[currentListKey] = [];
          meta[currentListKey].push(item);
        }
      }
      i++;
      continue;
    }
    currentListKey = null;
    if (line.includes(":")) {
      const idx = line.indexOf(":");
      const key = line.slice(0, idx).trim();
      let value = line.slice(idx + 1).trim();
      if (!key) {
        i++;
        continue;
      }
      if (value === "") {
        meta[key] = [];
        currentListKey = key;
        i++;
        continue;
      }
      if (value.startsWith("[") && value.endsWith("]")) {
        const inner = value.slice(1, -1).trim();
        meta[key] = inner
          ? inner
              .split(",")
              .map((p) => p.trim().replace(/^["']|["']$/g, ""))
              .filter(Boolean)
          : [];
        i++;
        continue;
      }
      meta[key] = value.replace(/^["']|["']$/g, "");
    }
    i++;
  }
  return { meta, rest: lines.slice(i).join("\n") };
}

function parseControlledMarkdown(text) {
  const { meta, rest: bodyText } = parseSimpleFrontMatter(text);

  const blocks = [];
  for (const rawLine of bodyText.split("\n")) {
    const line = rawLine.trimEnd();
    if (!line.trim()) continue;
    if (line.startsWith("#")) {
      let level = 0;
      for (const ch of line) {
        if (ch === "#") level++;
        else break;
      }
      if (level >= 1 && level <= 5) {
        const content = line.slice(level).trimStart();
        if (content) {
          blocks.push({ type: `h${level}`, text: content });
          continue;
        }
      }
    }
    blocks.push({ type: "p", text: line.trim() });
  }

  let title = "";
  for (const b of blocks) {
    if (b.type === "h1") {
      title = b.text.trim();
      break;
    }
  }

  return {
    title,
    doc_number: meta.doc_number || null,
    recipients: meta.recipients || null,
    attachments: meta.attachments || null,
    signer: meta.signer || null,
    date: meta.date || null,
    contact: meta.contact || null,
    blocks,
  };
}

function chooseFontForParagraph(text) {
  const stripped = text.trim();
  const cjkNumerals = "\u4E00\u4E8C\u4E09\u56DB\u4E94\u516D\u4E03\u516B\u4E5D\u5341";
  if (cjkNumerals.includes(stripped[0]) && stripped.includes("\u3001") && stripped.indexOf("\u3001") <= 3) {
    return FONT_LEVEL1;
  }
  if (stripped.startsWith("\uFF08") && stripped.includes("\uFF09") && stripped.indexOf("\uFF09") <= 4) {
    return FONT_BODY;
  }
  return FONT_BODY;
}

function makeRun(text, fontCN, sizePt, options = {}) {
  const { bold = false, asciiFont } = options;
  const runProps = {
    text: normalizeQuotes(text),
    font: {
      ascii: asciiFont || fontCN,
      hAnsi: asciiFont || fontCN,
      eastAsia: fontCN,
    },
    size: ptToHalfPt(sizePt),
    bold,
  };
  return new TextRun(runProps);
}

function makeParagraph(text, fontCN, sizePt, options = {}) {
  const {
    align = AlignmentType.JUSTIFIED,
    firstLineIndent = null,
    rightIndent = null,
    asciiFont,
    bold = false,
  } = options;

  const paraProps = {
    children: [makeRun(text, fontCN, sizePt, { bold, asciiFont })],
    spacing: {
      before: 0,
      after: 0,
      line: LINE_SPACING,
      lineRule: "exact",
    },
    alignment: align,
  };

  if (firstLineIndent !== null) {
    paraProps.indent = { firstLine: firstLineIndent };
  }
  if (rightIndent !== null) {
    paraProps.indent = paraProps.indent || {};
    paraProps.indent.right = rightIndent;
  }

  return new Paragraph(paraProps);
}

function makeBlankParagraph() {
  return new Paragraph({
    children: [],
    spacing: {
      before: 0,
      after: 0,
      line: LINE_SPACING,
      lineRule: "exact",
    },
  });
}

function addDocNumber(children, docNumber) {
  if (!docNumber) return;
  children.push(makeBlankParagraph());
  children.push(
    makeParagraph(docNumber, FONT_BODY, SIZE_BODY, {
      align: AlignmentType.RIGHT,
      asciiFont: FONT_ASCII,
    })
  );
  children.push(makeBlankParagraph());
}

function addDisclosureParagraph(children) {
  children.push(
    new Paragraph({
      children: [
        makeRun("\u516C\u5F00\u65B9\u5F0F\uFF1A", FONT_DISCLOSURE_KEY, SIZE_DISCLOSURE),
        makeRun("\u4F9D\u7533\u8BF7\u516C\u5F00", FONT_DISCLOSURE_VALUE, SIZE_DISCLOSURE),
      ],
      spacing: {
        before: 0,
        after: 0,
        line: LINE_SPACING,
        lineRule: "exact",
      },
      alignment: AlignmentType.JUSTIFIED,
      indent: { firstLine: TWO_CHAR_INDENT },
    })
  );
}

function hasDisclosureInChildren(children) {
  for (const child of children) {
    if (child instanceof Paragraph) {
      const runs = child.root
        ? JSON.stringify(child.root)
        : "";
      if (runs.includes("\u516C\u5F00\u65B9\u5F0F") && runs.includes("\u4F9D\u7533\u8BF7\u516C\u5F00")) {
        return true;
      }
    }
  }
  return false;
}

function estimateTextWidthHanzi(text) {
  if (!text) return 0;
  let width = 0;
  for (const ch of text) {
    if (/\s/.test(ch)) continue;
    const code = ch.codePointAt(0);
    const isWide =
      (code >= 0x1100 && code <= 0x115f) ||
      (code >= 0x2e80 && code <= 0x303f) ||
      (code >= 0x3040 && code <= 0x33ff) ||
      (code >= 0x3400 && code <= 0x4dbf) ||
      (code >= 0x4e00 && code <= 0x9fff) ||
      (code >= 0xac00 && code <= 0xd7a3) ||
      (code >= 0xf900 && code <= 0xfaff) ||
      (code >= 0xfe30 && code <= 0xfe6f) ||
      (code >= 0xff01 && code <= 0xff60) ||
      (code >= 0xffe0 && code <= 0xffe6) ||
      (code >= 0x20000 && code <= 0x2fffd) ||
      (code >= 0x30000 && code <= 0x3fffd);
    width += isWide ? 1.0 : 0.5;
  }
  return width;
}

const DATE_YMD_RE = /^\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*$/;

function toAsciiDigits(text) {
  if (!text) return text;
  return text.replace(/[\uff10-\uff19]/g, (ch) =>
    String.fromCharCode(ch.charCodeAt(0) - 0xfee0)
  );
}

function estimateDateWidthHanzi(dateText) {
  if (!dateText) return null;
  const normalized = toAsciiDigits(dateText.trim());
  const m = normalized.match(DATE_YMD_RE);
  if (!m) return null;
  const month2 = m[2].length === 2;
  const day2 = m[3].length === 2;
  if (!month2 && !day2) return 7.5;
  if (month2 && day2) return 8.2;
  return 7.8;
}

function computeDateRightIndent(signer, date) {
  signer = (signer || "").trim();
  date = (date || "").trim();
  if (!signer || !date) return RIGHT_INDENT_DATE;

  const signerWidth = estimateTextWidthHanzi(signer);
  const dateWidth = estimateDateWidthHanzi(date);
  const dateWidthSource = dateWidth !== null ? "mapping" : "fallback";
  const effectiveDateWidth = dateWidth !== null ? dateWidth : estimateTextWidthHanzi(date);

  const twipsPerHanzi = RIGHT_INDENT_SIGNER / 4.0;

  const indent =
    RIGHT_INDENT_SIGNER +
    (signerWidth - effectiveDateWidth) * 0.5 * twipsPerHanzi;
  return Math.max(0, Math.round(indent));
}

function extractTopLevelParas(xml, maxCount) {
  const bodyIdx = xml.indexOf("<w:body");
  if (bodyIdx === -1) return [];
  let pos = xml.indexOf(">", bodyIdx) + 1;
  const openRe = /<w:p(?:\s[^>]*>|>|\/>)/g;
  const closeTag = "</w:p>";
  const result = [];
  let depth = 0, paraStart = -1, found = 0;

  while (found < maxCount && pos < xml.length) {
    openRe.lastIndex = pos;
    const om = openRe.exec(xml);
    const oi = om ? om.index : -1;
    const ci = xml.indexOf(closeTag, pos);
    if (oi === -1 && ci === -1) break;

    if (oi !== -1 && (ci === -1 || oi < ci)) {
      const ts = om[0];
      if (ts.endsWith("/>")) {
        if (depth === 0) { result.push(ts); found++; }
        pos = oi + ts.length;
      } else {
        if (depth === 0) paraStart = oi;
        depth++;
        pos = oi + ts.length;
      }
    } else {
      depth--;
      if (depth === 0 && paraStart !== -1) {
        result.push(xml.substring(paraStart, ci + closeTag.length));
        found++;
        paraStart = -1;
      }
      pos = ci + closeTag.length;
    }
  }
  return result;
}

function mergeNamespaceDecls(targetXml, sourceXml) {
  const nsRe = /xmlns:(\w+)="([^"]+)"/g;
  const srcRoot = sourceXml.match(/<w:document[^>]*>/);
  if (!srcRoot) return targetXml;
  const srcNs = {};
  let m;
  while ((m = nsRe.exec(srcRoot[0])) !== null) srcNs[m[1]] = m[2];

  const tgtRoot = targetXml.match(/<w:document[^>]*>/);
  if (!tgtRoot) return targetXml;
  const tgtNs = {};
  nsRe.lastIndex = 0;
  while ((m = nsRe.exec(tgtRoot[0])) !== null) tgtNs[m[1]] = m[2];

  const missing = [];
  for (const [p, u] of Object.entries(srcNs)) {
    if (!(p in tgtNs)) missing.push(`xmlns:${p}="${u}"`);
  }
  if (!missing.length) return targetXml;

  const insertAt = targetXml.indexOf(">", tgtRoot.index);
  return targetXml.slice(0, insertAt) + " " + missing.join(" ") + targetXml.slice(insertAt);
}

async function buildWithTemplate(doc, templatePath, outputPath) {
  const contentBuffer = await Packer.toBuffer(doc);
  const templateData = fs.readFileSync(templatePath);
  const tZip = await JSZip.loadAsync(templateData);
  const cZip = await JSZip.loadAsync(contentBuffer);

  const tXmlFile = tZip.file("word/document.xml");
  if (!tXmlFile) {
    fs.writeFileSync(outputPath, contentBuffer);
    console.log(`已生成：${outputPath}`);
    return;
  }
  const tXml = await tXmlFile.async("string");
  const redHeadParas = extractTopLevelParas(tXml, 2);
  if (redHeadParas.length < 2) {
    fs.writeFileSync(outputPath, contentBuffer);
    console.log(`已生成：${outputPath}`);
    return;
  }

  let cXml = await cZip.file("word/document.xml").async("string");
  cXml = mergeNamespaceDecls(cXml, tXml);

  // 提取生成文档的 body 内部内容（去掉 sectPr）
  const cBodyMatch = cXml.match(/<w:body[^>]*>([\s\S]*)<\/w:body>/);
  let cBodyInner = cBodyMatch ? cBodyMatch[1] : "";
  cBodyInner = cBodyInner.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, "");

  // 提取模板的 sectPr（完整保留页眉/页脚引用）
  const tSectPrMatch = tXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/);
  const tSectPr = tSectPrMatch ? tSectPrMatch[0] : "";

  // 组装新的 body：红头 + 生成正文 + 模板 sectPr
  const newBodyInner = redHeadParas.join("") + cBodyInner + tSectPr;

  const bodyOpenMatch = cXml.match(/<w:body[^>]*>/);
  const bodyCloseMatch = cXml.match(/<\/w:body>/);
  if (bodyOpenMatch && bodyCloseMatch) {
    const start = cXml.indexOf(bodyOpenMatch[0]) + bodyOpenMatch[0].length;
    const end = cXml.indexOf(bodyCloseMatch[0]);
    cXml = cXml.slice(0, start) + newBodyInner + cXml.slice(end);
  }

  // 以模板为基底，替换 document.xml
  tZip.file("word/document.xml", cXml);

  // 合并 document.xml.rels：把生成文档中有但模板中没有的唯一 ID 关系添加进去
  const tRelsFile = tZip.file("word/_rels/document.xml.rels");
  const cRelsFile = cZip.file("word/_rels/document.xml.rels");
  if (tRelsFile && cRelsFile) {
    const tRels = await tRelsFile.async("string");
    const cRels = await cRelsFile.async("string");

    // 收集模板中已有的关系 ID，避免重复
    const existingIds = new Set();
    const idRe = /Id="(rId\d+)"/g;
    let im;
    while ((im = idRe.exec(tRels)) !== null) existingIds.add(im[1]);

    // 从生成文档中提取关系，只添加 ID 不冲突的
    const relRe = /<Relationship\s[^>]*\/>/g;
    const newEntries = [];
    let rm;
    while ((rm = relRe.exec(cRels)) !== null) {
      const entry = rm[0];
      const idMatch = entry.match(/Id="(rId\d+)"/);
      if (idMatch && !existingIds.has(idMatch[1])) {
        newEntries.push(entry);
        existingIds.add(idMatch[1]);
      }
    }

    if (newEntries.length) {
      const closePos = tRels.indexOf("</Relationships>");
      if (closePos !== -1) {
        tZip.file("word/_rels/document.xml.rels",
          tRels.slice(0, closePos) + newEntries.join("\n") + tRels.slice(closePos));
      }
    }
  }

  // 复制生成文档中独有的 word/ 文件到模板（字体、styles.xml 等）
  const cKeys = Object.keys(cZip.files);
  for (const f of cKeys) {
    if (f.startsWith("word/") && !f.endsWith("/") && !tZip.file(f)) {
      tZip.file(f, await cZip.files[f].async("arraybuffer"));
    }
  }

  const out = await tZip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(outputPath, out);
  console.log(`已生成（含红头）：${outputPath}`);
}

function estimatePageCount(docXml) {
  const bodyMatch = docXml.match(/<w:body[^>]*>([\s\S]*)<\/w:body>/);
  if (!bodyMatch) return 1;
  const bodyInner = bodyMatch[1];
  const paraCount = (bodyInner.match(/<w:p[\s>]/g) || []).length;
  const LINES_PER_PAGE = 23;
  return Math.max(1, Math.ceil(paraCount / LINES_PER_PAGE));
}

async function injectBanjiSection(outputPath) {
  const data = fs.readFileSync(outputPath);
  const zip = await JSZip.loadAsync(data);

  // 找到最大 rId
  const relsXml = await zip.file("word/_rels/document.xml.rels").async("string");
  let maxId = 0;
  for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) {
    maxId = Math.max(maxId, parseInt(m[1]));
  }
  const newRId = "rId" + (maxId + 1);

  // 创建版记 footer 文件
  const banjiPara = `<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="28"/></w:rPr><w:t>公开方式：</w:t></w:r><w:r><w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">依申请公开</w:t></w:r></w:p>`;

  // 从模板默认页脚提取原始 <w:ftr> 命名空间声明和页码内容
  let footerXml;
  const templateFooter = zip.file("word/footer2.xml");
  if (templateFooter) {
    const tf = await templateFooter.async("string");
    const rootMatch = tf.match(/^(\s*<\?xml[^?]*\?>\s*)?(<w:ftr[^>]*>)/);
    if (rootMatch) {
      const ftrOpen = rootMatch[2];
      const contentStart = rootMatch[0].length;
      const contentEnd = tf.lastIndexOf('</w:ftr>');
      const innerContent = tf.slice(contentStart, contentEnd).replace(
        /w:jc w:val="right"/g,
        'w:jc w:val="left"'
      );
      footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${ftrOpen}${banjiPara}${innerContent}</w:ftr>`;
    } else {
      footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${banjiPara}</w:ftr>`;
    }
  } else {
    footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${banjiPara}</w:ftr>`;
  }
  zip.file("word/footer99.xml", footerXml);

  // 注册 footer99.xml 到 [Content_Types].xml（Word 严格校验此项）
  const ctXml = await zip.file("[Content_Types].xml").async("string");
  const ctOverride = '<Override PartName="/word/footer99.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';
  if (!ctXml.includes("footer99.xml")) {
    zip.file("[Content_Types].xml", ctXml.replace("</Types>", ctOverride + "</Types>"));
  }

  // 添加 rels 关系
  const footerRel = `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer99.xml"/>`;
  const relsUpdated = relsXml.replace("</Relationships>", footerRel + "</Relationships>");
  zip.file("word/_rels/document.xml.rels", relsUpdated);

  // 修改 document.xml：先插入临时版记行估页数，再包进 section break
  let docXml = await zip.file("word/document.xml").async("string");
  const sectPrMatch = docXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/);
  if (sectPrMatch) {
    const origSectPr = sectPrMatch[0];
    // 临时版记段落：格式与最终版记一致，插入正文末用于精确估计页数
    const tempBanjiPara = `<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="540" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="黑体"/><w:sz w:val="28"/></w:rPr><w:t>公开方式：</w:t></w:r><w:r><w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">依申请公开</w:t></w:r></w:p>`;
    const docXmlWithBanji = docXml.replace(origSectPr, tempBanjiPara + origSectPr);
    const pageCount = estimatePageCount(docXmlWithBanji);
    let modSectPr = origSectPr;
    if (pageCount % 2 === 0) {
      modSectPr = origSectPr.replace(/(<w:sectPr[^>]*>)/, '$1<w:type w:val="continuous"/>');
    }
    const sectionBreak = `<w:p><w:pPr>${modSectPr}</w:pPr></w:p>`;
    // 空段落（让 section 2 有内容，防止被忽略）
    const emptyPara = `<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>`;
    // Section 2 的 sectPr（版记 footer，显式设置所有 footer 类型防止继承）
    const sec2SectPr = `<w:sectPr><w:footerReference w:type="even" r:id="${newRId}"/><w:footerReference w:type="default" r:id="${newRId}"/><w:footerReference w:type="first" r:id="${newRId}"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2098" w:right="1588" w:bottom="2041" w:left="1588" w:header="851" w:footer="1644" w:gutter="0"/></w:sectPr>`;
    docXml = docXml.replace(origSectPr, sectionBreak + emptyPara + sec2SectPr);
    zip.file("word/document.xml", docXml);
  }

  const out = await zip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(outputPath, out);
}

function buildDocument(data) {
  const title = String(data.title || "").trim();
  const docNumber = String(data.doc_number || "").trim();
  const recipients = normalizeRecipients(data.recipients);
  let bodyList = data.body || [];
  const attachments = normalizeList(data.attachments);
  const signer = String(data.signer || "").trim();
  const date = String(data.date || "").trim();
  const contact = String(data.contact || "").trim();
  const blocks = data.blocks;

  if (typeof bodyList === "string") {
    bodyList = parseBodyText(bodyList);
  } else if (Array.isArray(bodyList)) {
    bodyList = bodyList.map((s) => String(s).trim()).filter(Boolean);
  } else {
    bodyList = [];
  }

  const children = [];

  if (Array.isArray(blocks) && blocks.length > 0) {
    let titleRendered = false;

    addDocNumber(children, docNumber);

    for (const block of blocks) {
      if (block.type === "h1") {
        const text = String(block.text || "").trim();
        if (text) {
          children.push(
            makeParagraph(text, FONT_HEADER, SIZE_TITLE, {
              align: AlignmentType.CENTER,
            })
          );
          children.push(makeBlankParagraph());
          titleRendered = true;
        }
        break;
      }
    }

    if (!titleRendered && title) {
      children.push(
        makeParagraph(title, FONT_HEADER, SIZE_TITLE, {
          align: AlignmentType.CENTER,
        })
      );
      children.push(makeBlankParagraph());
    }

    if (recipients) {
      children.push(
        makeParagraph(`${recipients}\uFF1A`, FONT_BODY, SIZE_BODY, {
          align: AlignmentType.LEFT,
          asciiFont: FONT_ASCII,
        })
      );
    }

    let skippedFirstH1 = false;
    for (const block of blocks) {
      const btype = block.type;
      const text = String(block.text || "").trim();
      if (!text) continue;

      if (btype === "h1" && !skippedFirstH1) {
        skippedFirstH1 = true;
        continue;
      }

      if (btype && btype.startsWith("h")) {
        const level = parseInt(btype.slice(1), 10);
        if (level === 2) {
          children.push(
            makeParagraph(text, FONT_LEVEL1, SIZE_BODY, {
              firstLineIndent: TWO_CHAR_INDENT,
              asciiFont: FONT_ASCII,
            })
          );
        } else if (level >= 3) {
          children.push(
            makeParagraph(text, FONT_BODY, SIZE_BODY, {
              firstLineIndent: TWO_CHAR_INDENT,
              asciiFont: FONT_ASCII,
            })
          );
        } else {
          children.push(
            makeParagraph(text, FONT_BODY, SIZE_BODY, {
              firstLineIndent: TWO_CHAR_INDENT,
              asciiFont: FONT_ASCII,
            })
          );
        }
      } else {
        children.push(
          makeParagraph(text, FONT_BODY, SIZE_BODY, {
            firstLineIndent: TWO_CHAR_INDENT,
            asciiFont: FONT_ASCII,
          })
        );
      }
    }
  } else {
    addDocNumber(children, docNumber);

    if (title) {
      children.push(
        makeParagraph(title, FONT_HEADER, SIZE_TITLE, {
          align: AlignmentType.CENTER,
        })
      );
      children.push(makeBlankParagraph());
    }

    if (recipients) {
      children.push(
        makeParagraph(`${recipients}\uFF1A`, FONT_BODY, SIZE_BODY, {
          align: AlignmentType.LEFT,
          asciiFont: FONT_ASCII,
        })
      );
    }

    for (const paragraphText of bodyList) {
      const fontName = chooseFontForParagraph(paragraphText);
      const align =
        fontName === FONT_BODY ? AlignmentType.JUSTIFIED : AlignmentType.LEFT;
      children.push(
        makeParagraph(paragraphText, fontName, SIZE_BODY, {
          align,
          firstLineIndent: TWO_CHAR_INDENT,
          asciiFont: FONT_ASCII,
        })
      );
    }
  }

  if (attachments.length > 0) {
    children.push(makeBlankParagraph());
    if (attachments.length === 1) {
      children.push(
        makeParagraph(`\u9644\u4EF6\uFF1A${attachments[0]}`, FONT_BODY, SIZE_BODY, {
          firstLineIndent: TWO_CHAR_INDENT,
        })
      );
    } else {
      for (let idx = 0; idx < attachments.length; idx++) {
        const prefix = idx === 0 ? "\u9644\u4EF6\uFF1A" : "";
        const text = prefix
          ? `${prefix}${idx + 1}. ${attachments[idx]}`
          : `${idx + 1}. ${attachments[idx]}`;
        children.push(
          makeParagraph(text, FONT_BODY, SIZE_BODY, {
            firstLineIndent: TWO_CHAR_INDENT,
          })
        );
      }
    }
  }

  if (signer || date || contact) {
    children.push(makeBlankParagraph());
    if (signer) {
      children.push(
        makeParagraph(signer, FONT_BODY, SIZE_BODY, {
          align: AlignmentType.RIGHT,
          rightIndent: RIGHT_INDENT_SIGNER,
        })
      );
    }
    if (date) {
      const dateRight = computeDateRightIndent(signer, date);
      children.push(
        makeParagraph(date, FONT_BODY, SIZE_BODY, {
          align: AlignmentType.RIGHT,
          rightIndent: dateRight,
        })
      );
    }
    if (contact) {
      children.push(makeBlankParagraph());
      children.push(
        makeParagraph(
          `\uFF08\u8054\u7CFB\u4EBA\uFF1A${contact}\uFF09`,
          FONT_BODY,
          SIZE_BODY,
          {
            firstLineIndent: TWO_CHAR_INDENT,
          }
        )
      );
    }
  }

  

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              width: PAGE_WIDTH,
              height: PAGE_HEIGHT,
            },
            margin: {
              top: MARGIN_TOP,
              bottom: MARGIN_BOTTOM,
              left: MARGIN_LEFT,
              right: MARGIN_RIGHT,
            },
          },
        },
        children,
      },
    ],
  });

  return doc;
}

function main() {
  const args = process.argv.slice(2);
  let mdPath = null;
  let inputPath = null;
  let outputPath = "\u793A\u4F8B\u516C\u6587.docx";
  let cliData = {};

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case "--md":
        mdPath = args[++i];
        break;
      case "--input":
        inputPath = args[++i];
        break;
      case "-o":
      case "--output":
        outputPath = args[++i];
        break;
      case "--title":
        cliData.title = args[++i];
        break;
      case "--doc-number":
        cliData.doc_number = args[++i];
        break;
      case "--recipients":
        cliData.recipients = args[++i];
        break;
      case "--body":
        cliData.body = args[++i];
        break;
      case "--attachment":
        if (!cliData.attachments) cliData.attachments = [];
        cliData.attachments.push(args[++i]);
        break;
      case "--signer":
        cliData.signer = args[++i];
        break;
      case "--date":
        cliData.date = args[++i];
        break;
      case "--contact":
        cliData.contact = args[++i];
        break;
    }
  }

  let data = {};
  const hasCustom = Object.keys(cliData).length > 0;

  if (mdPath) {
    const content = fs.readFileSync(mdPath, "utf-8");
    data = parseControlledMarkdown(content);
  } else if (inputPath) {
    const ext = path.extname(inputPath).toLowerCase();
    if (ext === ".md" || ext === ".markdown") {
      const content = fs.readFileSync(inputPath, "utf-8");
      data = parseControlledMarkdown(content);
    } else {
      const content = fs.readFileSync(inputPath, "utf-8");
      data = JSON.parse(content);
    }
  } else if (!hasCustom) {
    data = {
      title: "\u5173\u4E8E\u5F00\u5C55\u5E74\u5EA6\u5DE5\u4F5C\u603B\u7ED3\u7684\u901A\u77E5",
      recipients: ["\u5404\u76F8\u5173\u5355\u4F4D"],
      body: [
        "\u4E3A\u5168\u9762\u603B\u7ED3\u5E74\u5EA6\u5DE5\u4F5C\u6210\u679C\uFF0C\u68B3\u7406\u7ECF\u9A8C\u505A\u6CD5\uFF0C\u73B0\u5C31\u6709\u5173\u4E8B\u9879\u901A\u77E5\u5982\u4E0B\u3002",
        "\u4E00\u3001\u603B\u4F53\u8981\u6C42",
        "\uFF08\u4E00\uFF09\u7A81\u51FA\u91CD\u70B9\u3002\u5404\u5355\u4F4D\u8981\u56F4\u7ED5\u4E2D\u5FC3\u4EFB\u52A1\uFF0C\u7A81\u51FA\u4EAE\u70B9\u5DE5\u4F5C\u3002",
        "1. \u505A\u5230\u6570\u636E\u51C6\u786E\u3001\u6750\u6599\u5B8C\u6574\u3002",
        "\uFF081\uFF09\u6309\u65F6\u62A5\u9001\uFF0C\u603B\u7ED3\u6750\u6599\u4E0D\u8D85\u8FC7\u4E09\u9875\u3002",
        "\u8BF7\u4E8E2\u670815\u65E5\u524D\u62A5\u9001\u7535\u5B50\u7248\u6750\u6599\u3002",
      ],
      attachments: ["\u5E74\u5EA6\u5DE5\u4F5C\u603B\u7ED3\u6A21\u677F"],
    };
  }

  data = { ...data, ...cliData };

  if (data.body && typeof data.body === "string") {
    data.body = parseBodyText(data.body);
  }

  const doc = buildDocument(data);

  if (!fs.existsSync(DEFAULT_TEMPLATE_PATH)) {
    console.error(`错误：红头模板不存在：${DEFAULT_TEMPLATE_PATH}`);
    process.exit(1);
  }

  buildWithTemplate(doc, DEFAULT_TEMPLATE_PATH, outputPath).then(() => injectBanjiSection(outputPath)).catch(console.error);
}

main();
