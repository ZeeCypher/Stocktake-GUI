const crypto = require("node:crypto");
const fs = require("node:fs");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");

const port = Number(process.env.PORT || 3000);
const appPin = String(process.env.STOCKTAKE_PIN || "1234");
const managerPin = String(process.env.STOCKTAKE_MANAGER_PIN || "4321");
const sessionCookieName = "stocktake_session";
const managerSessionCookieName = "stocktake_manager_session";
const indexFile = path.join(__dirname, "index.html");
const databaseFile = path.join(__dirname, "stocktake-data.json");
const sparkLogoFile = path.join(__dirname, "Spark_Logo_RGB.png");
const sessions = new Set();
const managerSessions = new Set();
const staticContentTypes = {
  ".gif": "image/gif",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".webp": "image/webp"
};

const uid = () => {
  if (typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }

  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
};

const now = () => new Date().toISOString();

const loginPage = (errorMessage = "") => `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Stocktake Login</title>
  <style>
    * { box-sizing: border-box; }
    body {
      min-height: 100vh;
      margin: 0;
      display: grid;
      place-items: center;
      color: #152033;
      background:
        linear-gradient(135deg, rgba(37, 99, 235, 0.11), transparent 34%),
        linear-gradient(315deg, rgba(14, 165, 233, 0.1), transparent 30%),
        #f3f7fc;
      font-family: "Segoe UI", Arial, sans-serif;
    }
    form {
      width: min(380px, calc(100vw - 32px));
      padding: 26px;
      background: #ffffff;
      border: 1px solid #d8e3f1;
      border-radius: 10px;
      box-shadow: 0 18px 40px rgba(21, 32, 51, 0.12);
    }
    img {
      display: block;
      width: 190px;
      max-height: 90px;
      object-fit: contain;
      margin-bottom: 18px;
    }
    h1 {
      margin: 0 0 8px;
      font-size: 24px;
      line-height: 1.1;
    }
    p {
      margin: 0 0 18px;
      color: #627086;
      font-size: 14px;
      line-height: 1.4;
    }
    label {
      display: block;
      margin-bottom: 7px;
      color: #627086;
      font-size: 11px;
      font-weight: 800;
      text-transform: uppercase;
    }
    input {
      width: 100%;
      min-height: 42px;
      padding: 10px 12px;
      border: 1px solid #b9c9e2;
      border-radius: 7px;
      font: inherit;
      outline: 0;
    }
    input:focus {
      border-color: #2563eb;
      box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.14);
    }
    button {
      width: 100%;
      min-height: 42px;
      margin-top: 14px;
      color: white;
      background: #2563eb;
      border: 0;
      border-radius: 7px;
      font: inherit;
      font-weight: 800;
      cursor: pointer;
    }
    .error {
      margin-top: 12px;
      color: #b63832;
      font-size: 13px;
      font-weight: 700;
    }
  </style>
</head>
<body>
  <form method="post" action="/api/login">
    <img src="/logo.png" alt="Company logo">
    <h1>Stocktake locked</h1>
    <p>Enter the PIN to open the shared stocktake app.</p>
    <label for="pin">PIN</label>
    <input id="pin" name="pin" type="password" inputmode="numeric" autocomplete="current-password" autofocus>
    <button type="submit">Unlock</button>
    ${errorMessage ? `<div class="error">${errorMessage}</div>` : ""}
  </form>
</body>
</html>`;

const defaultDatabase = () => ({
  items: [],
  deletedItems: [],
  logEntries: []
});

const cleanText = (value) => String(value || "").trim();

const cleanActor = (value) => cleanText(value);

const normalizeIp = (value) => {
  const ip = cleanText(value).split(",")[0].trim();
  if (!ip) return "";
  if (ip === "::1") return "127.0.0.1";
  if (ip.startsWith("::ffff:")) return ip.slice(7);
  return ip;
};

const requestIp = (request) =>
  normalizeIp(request.headers["x-forwarded-for"] || request.socket.remoteAddress || "");

const clientInfoForRequest = (request) => ({
  ipAddress: requestIp(request)
});


const cleanNumber = (value) => {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? Math.floor(number) : 0;
};

const statusFor = (item) => {
  if (item.retired) return "Retired";
  if (item.quantity <= 0) return "Out of stock";
  if (item.quantity <= item.minimum) return "Low stock";
  return "In stock";
};

const escapeXml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

const excelColumnName = (columnNumber) => {
  let name = "";
  let current = columnNumber;
  while (current > 0) {
    const mod = (current - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    current = Math.floor((current - mod) / 26);
  }
  return name;
};

const excelTextCell = (reference, value, style = 0) =>
  `<c r="${reference}" t="inlineStr" s="${style}"><is><t>${escapeXml(value)}</t></is></c>`;

const excelNumberCell = (reference, value, style = 0) =>
  `<c r="${reference}" s="${style}"><v>${Number.isFinite(Number(value)) ? Number(value) : 0}</v></c>`;

const excelRow = (rowNumber, cells, height = 0) => {
  const heightXml = height > 0 ? ` ht="${height}" customHeight="1"` : "";
  return `<row r="${rowNumber}"${heightXml}>${cells.join("")}</row>`;
};

const displayDate = (value) => {
  if (!value) return "Not updated";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return new Intl.DateTimeFormat("en-AU", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).format(date);
};

const reportSummaryForRows = (rows) => {
  const activeRows = rows.filter((row) => row.status !== "Retired");
  return {
    totalItems: activeRows.length,
    lowStock: activeRows.filter((row) => row.status === "Low stock").length,
    outOfStock: activeRows.filter((row) => row.status === "Out of stock").length,
    retired: rows.filter((row) => row.status === "Retired").length,
    categories: new Set(activeRows.map((row) => row.category).filter(Boolean)).size
  };
};

const normalizeExportRows = (rows) =>
  (Array.isArray(rows) ? rows : []).map((row) => ({
    item: cleanText(row.item || row.name),
    category: cleanText(row.category),
    location: cleanText(row.location),
    quantity: cleanNumber(row.quantity),
    minimum: cleanNumber(row.minimum),
    status: cleanText(row.status) || statusFor(row),
    notes: cleanText(row.notes),
    updated: cleanText(row.updated || row.updatedAt)
  }));

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let value = i;
    for (let bit = 0; bit < 8; bit++) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[i] = value >>> 0;
  }
  return table;
})();

const crc32 = (buffer) => {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
};

const createZip = (files) => {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  files.forEach((file) => {
    const nameBuffer = Buffer.from(file.name, "utf8");
    const data = Buffer.isBuffer(file.data) ? file.data : Buffer.from(String(file.data), "utf8");
    const crc = crc32(data);

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(data.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(nameBuffer.length, 26);
    localHeader.writeUInt16LE(0, 28);
    localParts.push(localHeader, nameBuffer, data);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(data.length, 20);
    centralHeader.writeUInt32LE(data.length, 24);
    centralHeader.writeUInt16LE(nameBuffer.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    centralParts.push(centralHeader, nameBuffer);

    offset += localHeader.length + nameBuffer.length + data.length;
  });

  const centralDirectory = Buffer.concat(centralParts);
  const localData = Buffer.concat(localParts);
  const endRecord = Buffer.alloc(22);
  endRecord.writeUInt32LE(0x06054b50, 0);
  endRecord.writeUInt16LE(0, 4);
  endRecord.writeUInt16LE(0, 6);
  endRecord.writeUInt16LE(files.length, 8);
  endRecord.writeUInt16LE(files.length, 10);
  endRecord.writeUInt32LE(centralDirectory.length, 12);
  endRecord.writeUInt32LE(localData.length, 16);
  endRecord.writeUInt16LE(0, 20);

  return Buffer.concat([localData, centralDirectory, endRecord]);
};

const buildStocktakeReportXlsx = (rows, reportDate = new Date()) => {
  const summary = reportSummaryForRows(rows);
  const createdAt = reportDate.toISOString().replace(/\.\d{3}Z$/, "Z");
  const hasLogo = fs.existsSync(sparkLogoFile);

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${hasLogo ? '<Default Extension="png" ContentType="image/png"/>' : ""}
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${hasLogo ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : ""}
</Types>`;

  const rootRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

  const workbookRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  const sheetRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;

  const drawingRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/Spark_Logo_RGB.png"/>
</Relationships>`;

  const drawing = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>6</xdr:col><xdr:colOff>120000</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>100000</xdr:rowOff></xdr:from>
    <xdr:ext cx="1050000" cy="795000"/>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="Spark IT Logo" descr="Spark IT Logo"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`;

  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Stocktake Report" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

  const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Spark IT Stocktake Report</dc:title><dc:creator>Spark IT Stocktake</dc:creator><cp:lastModifiedBy>Spark IT Stocktake</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${createdAt}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${createdAt}</dcterms:modified>
</cp:coreProperties>`;

  const app = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Spark IT Stocktake</Application></Properties>`;

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="7"><font><sz val="11"/><color rgb="FF14233D"/><name val="Segoe UI"/></font><font><b/><sz val="24"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font><font><sz val="11"/><color rgb="FFEAF2FF"/><name val="Segoe UI"/></font><font><b/><sz val="10"/><color rgb="FF24539A"/><name val="Segoe UI"/></font><font><b/><sz val="20"/><color rgb="FF14233D"/><name val="Segoe UI"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font><font><b/><sz val="10"/><color rgb="FF14233D"/><name val="Segoe UI"/></font></fonts>
  <fills count="10"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF103B77"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEFF6FF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF103B77"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE7F7ED"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF3DE"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFE9E9"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFECEFF5"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFE7ECF3"/></left><right style="thin"><color rgb="FFE7ECF3"/></right><top style="thin"><color rgb="FFE7ECF3"/></top><bottom style="thin"><color rgb="FFE7ECF3"/></bottom><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="14"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="4" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf><xf numFmtId="0" fontId="6" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="6" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="6" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="6" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf><xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  const summaryLabels = ["Total active items", "Low stock", "Out of stock", "Retired items", "Categories"];
  const summaryValues = [summary.totalItems, summary.lowStock, summary.outOfStock, summary.retired, summary.categories];
  const headers = ["Item", "Category", "Location", "Qty", "Min", "Status", "Notes", "Updated"];
  const sheetRows = [
    excelRow(1, [excelTextCell("A1", "Spark IT Stocktake Summary", 1)], 34),
    excelRow(2, [excelTextCell("A2", `Report exported on ${displayDate(reportDate.toISOString())}`, 2)], 24),
    excelRow(4, summaryLabels.map((label, index) => excelTextCell(`${excelColumnName(index + 1)}4`, label, 3)), 22),
    excelRow(5, summaryValues.map((value, index) => excelNumberCell(`${excelColumnName(index + 1)}5`, value, 4)), 30),
    excelRow(7, headers.map((header, index) => excelTextCell(`${excelColumnName(index + 1)}7`, header, 5)), 24)
  ];

  rows.forEach((row, index) => {
    const rowNumber = index + 8;
    const statusStyle = row.status === "Out of stock" ? 10 : row.status === "Low stock" ? 9 : row.status === "Retired" ? 11 : 8;
    sheetRows.push(excelRow(rowNumber, [
      excelTextCell(`A${rowNumber}`, row.item, 6),
      excelTextCell(`B${rowNumber}`, row.category, 6),
      excelTextCell(`C${rowNumber}`, row.location, 6),
      excelNumberCell(`D${rowNumber}`, row.quantity, 7),
      excelNumberCell(`E${rowNumber}`, row.minimum, 7),
      excelTextCell(`F${rowNumber}`, row.status, statusStyle),
      excelTextCell(`G${rowNumber}`, row.notes, 13),
      excelTextCell(`H${rowNumber}`, displayDate(row.updated), 12)
    ], 21));
  });

  const lastRow = Math.max(7, rows.length + 7);
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="7" topLeftCell="A8" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A8" sqref="A8"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols><col min="1" max="1" width="31" customWidth="1"/><col min="2" max="2" width="18" customWidth="1"/><col min="3" max="3" width="20" customWidth="1"/><col min="4" max="5" width="10" customWidth="1"/><col min="6" max="6" width="16" customWidth="1"/><col min="7" max="7" width="24" customWidth="1"/><col min="8" max="8" width="22" customWidth="1"/></cols>
  <sheetData>${sheetRows.join("\n")}</sheetData>
  <autoFilter ref="A7:H${lastRow}"/>
  <mergeCells count="2"><mergeCell ref="A1:F1"/><mergeCell ref="A2:F2"/></mergeCells>
  <pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>
  ${hasLogo ? '<drawing r:id="rId1"/>' : ""}
</worksheet>`;

  const files = [
    { name: "[Content_Types].xml", data: contentTypes },
    { name: "_rels/.rels", data: rootRelationships },
    { name: "docProps/core.xml", data: core },
    { name: "docProps/app.xml", data: app },
    { name: "xl/workbook.xml", data: workbook },
    { name: "xl/_rels/workbook.xml.rels", data: workbookRelationships },
    { name: "xl/styles.xml", data: styles },
    { name: "xl/worksheets/sheet1.xml", data: sheet }
  ];

  if (hasLogo) {
    files.push(
      { name: "xl/worksheets/_rels/sheet1.xml.rels", data: sheetRelationships },
      { name: "xl/drawings/drawing1.xml", data: drawing },
      { name: "xl/drawings/_rels/drawing1.xml.rels", data: drawingRelationships },
      { name: "xl/media/Spark_Logo_RGB.png", data: fs.readFileSync(sparkLogoFile) }
    );
  }

  return createZip(files);
};

const exportFilename = (date = new Date()) => {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `stocktake-report-${day}-${month}-${year}.xlsx`;
};

const readDatabase = () => {
  if (!fs.existsSync(databaseFile)) {
    const database = defaultDatabase();
    writeDatabase(database);
    return database;
  }

  const database = JSON.parse(fs.readFileSync(databaseFile, "utf8"));
  const normalized = {
    items: Array.isArray(database.items) ? database.items : [],
    deletedItems: Array.isArray(database.deletedItems) ? database.deletedItems : [],
    logEntries: Array.isArray(database.logEntries) ? database.logEntries : []
  };

  normalized.logEntries = normalized.logEntries.map((entry) => ({
    ...entry,
    actor: cleanText(entry.actor) || "Before user tracking",
    ipAddress: cleanText(entry.ipAddress)
  }));

  return normalized;
};

const writeDatabase = (database) => {
  fs.writeFileSync(databaseFile, JSON.stringify(database, null, 2));
};

const addLog = (database, title, detail, actor, source = {}) => {
  database.logEntries.unshift({
    id: uid(),
    title,
    detail,
    actor: cleanActor(actor),
    ipAddress: cleanText(source.ipAddress),
    at: now()
  });
  database.logEntries = database.logEntries.slice(0, 100);
};

const sanitizeItem = (item = {}) => ({
  id: cleanText(item.id),
  name: cleanText(item.name),
  category: cleanText(item.category),
  location: cleanText(item.location),
  quantity: cleanNumber(item.quantity),
  minimum: cleanNumber(item.minimum),
  notes: cleanText(item.notes),
  retired: Boolean(item.retired),
  updatedAt: now()
});

const handleAction = (body, isManager, requestSource = {}) => {
  let database = readDatabase();
  let message = "Saved.";
  const actor = cleanActor(body.actor);
  if (!actor) {
    throw new Error("Enter your name before changing stock.");
  }
  const logSource = {
    ipAddress: requestSource.ipAddress
  };
  const addActionLog = (title, detail) => addLog(database, title, detail, actor, logSource);

  if (body.type === "saveItem") {
    if (!isManager) {
      throw new Error("Manager PIN required to add or edit items.");
    }

    const incoming = sanitizeItem(body.item);
    if (!incoming.name || !incoming.category || !incoming.location) {
      throw new Error("Name, category, and location are required.");
    }

    const existingIndex = database.items.findIndex((item) => item.id === incoming.id);
    if (existingIndex >= 0) {
      const existing = database.items[existingIndex];
      incoming.id = existing.id;
      incoming.retired = Boolean(existing.retired);
      incoming.updatedBy = actor;
      database.items[existingIndex] = incoming;
      addActionLog(`Updated ${incoming.name}`, `Quantity ${existing.quantity} to ${incoming.quantity}`);
      message = "Item updated.";
    } else {
      incoming.id = uid();
      incoming.retired = false;
      incoming.updatedBy = actor;
      database.items.unshift(incoming);
      addActionLog(`Added ${incoming.name}`, `Quantity ${incoming.quantity}`);
      message = "Item added.";
    }
  } else if (body.type === "deleteItem") {
    if (!isManager) {
      throw new Error("Manager PIN required to delete items.");
    }

    const itemIndex = database.items.findIndex((entry) => entry.id === body.id);
    if (itemIndex < 0) {
      throw new Error("Item was not found.");
    }

    const [deletedItem] = database.items.splice(itemIndex, 1);
    database.deletedItems.unshift({
      ...deletedItem,
      deletedAt: now(),
      deletedBy: actor
    });
    addActionLog(`Moved ${deletedItem.name} to bin`, "Deleted item can be restored");
    message = "Item moved to bin.";
  } else if (body.type === "restoreDeletedItem") {
    if (!isManager) {
      throw new Error("Manager PIN required to restore deleted items.");
    }

    const itemIndex = database.deletedItems.findIndex((entry) => entry.id === body.id);
    if (itemIndex < 0) {
      throw new Error("Deleted item was not found.");
    }

    const [restoredItem] = database.deletedItems.splice(itemIndex, 1);
    delete restoredItem.deletedAt;
    delete restoredItem.deletedBy;
    restoredItem.updatedAt = now();
    restoredItem.updatedBy = actor;
    database.items.unshift(restoredItem);
    addActionLog(`Restored ${restoredItem.name} from bin`, statusFor(restoredItem));
    message = "Item restored from bin.";
  } else if (body.type === "moveItem") {
    if (!isManager) {
      throw new Error("Manager PIN required to reorder items.");
    }

    const fromIndex = database.items.findIndex((entry) => entry.id === body.id);
    const targetIndex = database.items.findIndex((entry) => entry.id === body.targetId);
    const position = body.position === "after" ? "after" : "before";

    if (fromIndex < 0 || targetIndex < 0) {
      throw new Error("Item was not found.");
    }

    if (fromIndex !== targetIndex) {
      const [movedItem] = database.items.splice(fromIndex, 1);
      const adjustedTargetIndex = database.items.findIndex((entry) => entry.id === body.targetId);
      const insertIndex = position === "after" ? adjustedTargetIndex + 1 : adjustedTargetIndex;
      database.items.splice(insertIndex, 0, movedItem);
      movedItem.updatedAt = now();
      movedItem.updatedBy = actor;
      addActionLog(`Moved ${movedItem.name}`, position === "after" ? "Moved down" : "Moved up");
    }

    message = "Item order updated.";
  } else if (body.type === "rowAction") {
    const item = database.items.find((entry) => entry.id === body.id);
    if (!item) {
      throw new Error("Item was not found.");
    }

    if (body.action === "retire") {
      item.retired = !item.retired;
      item.updatedAt = now();
      item.updatedBy = actor;
      addActionLog(`${item.retired ? "Retired" : "Restored"} ${item.name}`, statusFor(item));
      message = item.retired ? "Item retired." : "Item restored.";
    } else {
      if (item.retired) {
        throw new Error("Restore the item before changing quantity.");
      }

      const before = item.quantity;
      if (body.action === "minus") {
        item.quantity = Math.max(0, item.quantity - 1);
      } else if (body.action === "plus") {
        item.quantity += 1;
      } else {
        throw new Error("Unknown stock action.");
      }

      item.updatedAt = now();
      item.updatedBy = actor;
      addActionLog(`${body.action === "plus" ? "Added 1 to" : "Removed 1 from"} ${item.name}`, `${before} to ${item.quantity}`);
      message = "Quantity updated.";
    }
  } else {
    throw new Error("Unknown request.");
  }

  writeDatabase(database);
  return { message, state: database };
};

const readBody = (request) =>
  new Promise((resolve, reject) => {
    let body = "";

    request.on("data", (chunk) => {
      body += chunk;
      if (body.length > 1_000_000) {
        request.destroy();
        reject(new Error("Request is too large."));
      }
    });

    request.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        reject(new Error("Invalid JSON request."));
      }
    });
  });

const parseFormBody = (body) =>
  Object.fromEntries(new URLSearchParams(String(body || "")));

const parseCookies = (cookieHeader = "") =>
  Object.fromEntries(
    cookieHeader
      .split(";")
      .map((cookie) => cookie.trim())
      .filter(Boolean)
      .map((cookie) => {
        const equalsIndex = cookie.indexOf("=");
        if (equalsIndex < 0) {
          return [cookie, ""];
        }

        return [
          decodeURIComponent(cookie.slice(0, equalsIndex)),
          decodeURIComponent(cookie.slice(equalsIndex + 1))
        ];
      })
  );

const isAuthenticated = (request) => {
  const cookies = parseCookies(request.headers.cookie);
  return sessions.has(cookies[sessionCookieName]);
};

const isManagerAuthenticated = (request) => {
  const cookies = parseCookies(request.headers.cookie);
  return managerSessions.has(cookies[managerSessionCookieName]);
};

const sendJson = (response, statusCode, body) => {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(JSON.stringify(body));
};

const sendIndex = (response) => {
  response.writeHead(200, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(fs.readFileSync(indexFile));
};

const sendLogin = (response, statusCode = 200, errorMessage = "") => {
  response.writeHead(statusCode, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store"
  });
  response.end(loginPage(errorMessage));
};

const sendStaticFile = (response, urlPathname) => {
  const fileName = path.basename(decodeURIComponent(urlPathname));
  const filePath = path.join(__dirname, fileName);
  const extension = path.extname(filePath).toLowerCase();
  const contentType = staticContentTypes[extension];

  if (!contentType || !fs.existsSync(filePath)) {
    return false;
  }

  response.writeHead(200, {
    "Content-Type": contentType,
    "Cache-Control": "no-cache"
  });
  response.end(fs.readFileSync(filePath));
  return true;
};

const server = http.createServer(async (request, response) => {
  try {
    const url = new URL(request.url, `http://${request.headers.host}`);

    if (request.method === "GET" && sendStaticFile(response, url.pathname)) {
      return;
    }

    if (request.method === "GET" && (url.pathname === "/" || url.pathname === "/index.html")) {
      if (isAuthenticated(request)) {
        sendIndex(response);
      } else {
        sendLogin(response);
      }
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/login") {
      let body = "";
      request.on("data", (chunk) => {
        body += chunk;
      });
      request.on("end", () => {
        const fields = parseFormBody(body);
        if (String(fields.pin || "") !== appPin) {
          sendLogin(response, 401, "Incorrect PIN.");
          return;
        }

        const sessionId = uid();
        sessions.add(sessionId);
        response.writeHead(303, {
          "Location": "/",
          "Set-Cookie": `${sessionCookieName}=${encodeURIComponent(sessionId)}; HttpOnly; SameSite=Lax; Path=/; Max-Age=28800`,
          "Cache-Control": "no-store"
        });
        response.end();
      });
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/logout") {
      const cookies = parseCookies(request.headers.cookie);
      sessions.delete(cookies[sessionCookieName]);
      managerSessions.delete(cookies[managerSessionCookieName]);
      response.writeHead(303, {
        "Location": "/",
        "Set-Cookie": [
          `${sessionCookieName}=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0`,
          `${managerSessionCookieName}=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0`
        ],
        "Cache-Control": "no-store"
      });
      response.end();
      return;
    }

    if (!isAuthenticated(request)) {
      sendJson(response, 401, { error: "Enter the PIN to unlock Stocktake." });
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/manager-login") {
      const body = await readBody(request);
      if (String(body.pin || "") !== managerPin) {
        sendJson(response, 403, { error: "Incorrect manager PIN." });
        return;
      }

      const managerSessionId = uid();
      managerSessions.add(managerSessionId);
      response.writeHead(200, {
        "Content-Type": "application/json; charset=utf-8",
        "Set-Cookie": `${managerSessionCookieName}=${encodeURIComponent(managerSessionId)}; HttpOnly; SameSite=Lax; Path=/; Max-Age=28800`,
        "Cache-Control": "no-store"
      });
      response.end(JSON.stringify({
        message: "Manager tools unlocked.",
        state: {
          ...readDatabase(),
          managerUnlocked: true
        }
      }));
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/manager-logout") {
      const cookies = parseCookies(request.headers.cookie);
      managerSessions.delete(cookies[managerSessionCookieName]);
      response.writeHead(200, {
        "Content-Type": "application/json; charset=utf-8",
        "Set-Cookie": `${managerSessionCookieName}=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0`,
        "Cache-Control": "no-store"
      });
      response.end(JSON.stringify({
        message: "Manager tools locked.",
        state: {
          ...readDatabase(),
          managerUnlocked: false
        }
      }));
      return;
    }

    if (request.method === "GET" && url.pathname === "/api/state") {
      const clientInfo = clientInfoForRequest(request);
      sendJson(response, 200, {
        ...readDatabase(),
        managerUnlocked: isManagerAuthenticated(request),
        clientInfo
      });
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/export-report") {
      const body = await readBody(request);
      const database = readDatabase();
      const fallbackRows = database.items.map((item) => ({
        item: item.name,
        category: item.category,
        location: item.location,
        quantity: item.quantity,
        minimum: item.minimum,
        status: statusFor(item),
        notes: item.notes,
        updated: item.updatedAt
      }));
      const rows = normalizeExportRows(body.rows).filter((row) => row.item) || [];
      const reportRows = rows.length ? rows : normalizeExportRows(fallbackRows);
      const workbook = buildStocktakeReportXlsx(reportRows);
      const filename = exportFilename();

      response.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store"
      });
      response.end(workbook);
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/action") {
      const body = await readBody(request);
      const managerUnlocked = isManagerAuthenticated(request);
      const clientInfo = clientInfoForRequest(request);
      const result = handleAction(body, managerUnlocked, clientInfo);
      sendJson(response, 200, {
        ...result,
        state: {
          ...result.state,
          managerUnlocked,
          clientInfo
        }
      });
      return;
    }

    sendJson(response, 404, { error: "Not found." });
  } catch (error) {
    sendJson(response, 400, { error: error.message || "Server error." });
  }
});

server.on("error", (error) => {
  if (error.code === "EADDRINUSE") {
    console.error(`Port ${port} is already in use.`);
    console.error(`The stocktake app may already be running at http://localhost:${port}`);
    console.error(`Close the other server, or run this one on another port:`);
    console.error(`  $env:PORT=3001; node server.js`);
    process.exit(1);
  }

  throw error;
});

server.listen(port, "0.0.0.0", () => {
  console.log(`Stocktake database app is running:`);
  console.log(`  Local:   http://localhost:${port}`);

  Object.values(os.networkInterfaces())
    .flat()
    .filter((network) => network && network.family === "IPv4" && !network.internal)
    .forEach((network) => {
      console.log(`  Network: http://${network.address}:${port}`);
    });
});
