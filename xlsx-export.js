// Minimal, dependency-free .xlsx writer for Northern Projects exports.
// Builds a real Office Open XML workbook (stored, uncompressed) so Excel opens it
// natively with number formats, fills, borders, column widths and frozen headers.

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

const enc = (s) => new TextEncoder().encode(s);

// Build a ZIP archive with the STORE method — valid, and avoids a compression dependency.
function zip(files) {
  const chunks = [];
  const central = [];
  let offset = 0;

  const u16 = (v) => [v & 0xFF, (v >>> 8) & 0xFF];
  const u32 = (v) => [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF];

  for (const { name, data } of files) {
    const nameBytes = enc(name);
    const crc = crc32(data);
    const local = [
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0),
    ];
    chunks.push(new Uint8Array(local), nameBytes, data);
    central.push({ name: nameBytes, crc, size: data.length, offset });
    offset += local.length + nameBytes.length + data.length;
  }

  const cdChunks = [];
  let cdSize = 0;
  for (const e of central) {
    const rec = [
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(e.crc), ...u32(e.size), ...u32(e.size),
      ...u16(e.name.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0),
      ...u32(e.offset),
    ];
    cdChunks.push(new Uint8Array(rec), e.name);
    cdSize += rec.length + e.name.length;
  }

  const eocd = new Uint8Array([
    ...u32(0x06054b50), ...u16(0), ...u16(0),
    ...u16(central.length), ...u16(central.length),
    ...u32(cdSize), ...u32(offset), ...u16(0),
  ]);

  return new Blob([...chunks, ...cdChunks, eocd], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

const esc = (s) => String(s ?? '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/\u2212/g, '-');

function colName(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ---- Style catalogue -------------------------------------------------------
// Index order must match the XML written in styles() below.
const S = {
  base: 0, title: 1, subtitle: 2, label: 3, header: 4, section: 5,
  text: 6, money: 7, pct: 8, moneyBold: 9, total: 10, totalMoney: 11,
  good: 12, bad: 13, muted: 14, accent: 15, accentMoney: 16, date: 17,
};

function styles() {
  const fonts = [
    `<font><sz val="10"/><name val="Calibri"/></font>`,                                    // 0 base
    `<font><b/><sz val="16"/><color rgb="FF082058"/><name val="Calibri"/></font>`,          // 1 title
    `<font><sz val="10"/><color rgb="FF6B7488"/><name val="Calibri"/></font>`,              // 2 subtitle/muted
    `<font><b/><sz val="10"/><color rgb="FF082058"/><name val="Calibri"/></font>`,          // 3 label
    `<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>`,          // 4 header on navy
    `<font><b/><sz val="11"/><color rgb="FF082058"/><name val="Calibri"/></font>`,          // 5 section
    `<font><b/><sz val="10"/><name val="Calibri"/></font>`,                                 // 6 bold body
    `<font><sz val="10"/><color rgb="FF1F8A5B"/><name val="Calibri"/></font>`,              // 7 good
    `<font><sz val="10"/><color rgb="FFB8485B"/><name val="Calibri"/></font>`,              // 8 bad
  ];
  const fills = [
    `<fill><patternFill patternType="none"/></fill>`,
    `<fill><patternFill patternType="gray125"/></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FF082058"/></patternFill></fill>`, // 2 navy
    `<fill><patternFill patternType="solid"><fgColor rgb="FFEEF1F7"/></patternFill></fill>`, // 3 pale
    `<fill><patternFill patternType="solid"><fgColor rgb="FFF7F8FB"/></patternFill></fill>`, // 4 faint
  ];
  const borders = [
    `<border/>`,
    `<border><bottom style="thin"><color rgb="FFDFE4EC"/></bottom></border>`,               // 1 rule
    `<border><top style="medium"><color rgb="FF082058"/></top></border>`,                   // 2 total
  ];
  // numFmtId 164 = AED money, 165 = percent 2dp, 166 = date
  const numFmts = [
    `<numFmt numFmtId="164" formatCode="#,##0.00;[Red]-#,##0.00"/>`,
    `<numFmt numFmtId="165" formatCode="0.00%"/>`,
    `<numFmt numFmtId="166" formatCode="dd\\ mmm\\ yy"/>`,
  ];
  const xf = (o) =>
    `<xf numFmtId="${o.fmt || 0}" fontId="${o.font || 0}" fillId="${o.fill || 0}" borderId="${o.border || 0}" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">` +
    `<alignment horizontal="${o.align || 'general'}" vertical="center"${o.wrap ? ' wrapText="1"' : ''}/></xf>`;

  const cellXfs = [
    xf({}),                                                        // 0 base
    xf({ font: 1 }),                                               // 1 title
    xf({ font: 2 }),                                               // 2 subtitle
    xf({ font: 3 }),                                               // 3 label
    xf({ font: 4, fill: 2, align: 'center', wrap: true }),          // 4 header
    xf({ font: 5, fill: 3 }),                                      // 5 section
    xf({}),                                                        // 6 text
    xf({ fmt: 164, align: 'right' }),                              // 7 money
    xf({ fmt: 165, align: 'right' }),                              // 8 pct
    xf({ fmt: 164, font: 6, align: 'right' }),                     // 9 money bold
    xf({ font: 6, fill: 4, border: 2 }),                           // 10 total label
    xf({ fmt: 164, font: 6, fill: 4, border: 2, align: 'right' }),  // 11 total money
    xf({ fmt: 164, font: 7, align: 'right' }),                      // 12 good
    xf({ fmt: 164, font: 8, align: 'right' }),                      // 13 bad
    xf({ font: 2 }),                                               // 14 muted
    xf({ font: 4, fill: 2 }),                                      // 15 accent text
    xf({ fmt: 164, font: 4, fill: 2, align: 'right' }),             // 16 accent money
    xf({ fmt: 166, align: 'right' }),                              // 17 date
  ];

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>
<fonts count="${fonts.length}">${fonts.join('')}</fonts>
<fills count="${fills.length}">${fills.join('')}</fills>
<borders count="${borders.length}">${borders.join('')}</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${cellXfs.length}">${cellXfs.join('')}</cellXfs>
</styleSheet>`;
}

// A cell is either a primitive or { v, s, f } (value, style index, formula).
function cellXml(ref, cell) {
  if (cell === null || cell === undefined || cell === '') return '';
  const o = (typeof cell === 'object' && !(cell instanceof Date)) ? cell : { v: cell };
  const s = o.s !== undefined ? ` s="${o.s}"` : '';
  if (o.f) return `<c r="${ref}"${s}><f>${esc(o.f)}</f></c>`;
  if (typeof o.v === 'number' && isFinite(o.v)) return `<c r="${ref}"${s}><v>${o.v}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(o.v)}</t></is></c>`;
}

function sheetXml(sheet) {
  const rows = sheet.rows || [];
  const cols = (sheet.cols || [])
    .map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
  const body = rows.map((row, ri) => {
    if (!row || !row.length) return `<row r="${ri + 1}"/>`;
    const cells = row.map((c, ci) => cellXml(colName(ci + 1) + (ri + 1), c)).join('');
    return `<row r="${ri + 1}">${cells}</row>`;
  }).join('');
  const freeze = sheet.freeze
    ? `<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="${sheet.freeze}" topLeftCell="A${sheet.freeze + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`
    : `<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>`;
  const merges = (sheet.merges || []).length
    ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`
    : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${freeze}
${cols ? `<cols>${cols}</cols>` : ''}
<sheetData>${body}</sheetData>
${merges}
<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>
<pageSetup orientation="${sheet.landscape ? 'landscape' : 'portrait'}" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;
}

/**
 * Build an .xlsx Blob.
 * @param {Array<{name:string, rows:Array, cols?:number[], freeze?:number, merges?:string[], landscape?:boolean}>} sheets
 */
export function buildWorkbook(sheets) {
  const files = [];

  files.push({
    name: '[Content_Types].xml',
    data: enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}
</Types>`),
  });

  files.push({
    name: '_rels/.rels',
    data: enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
  });

  files.push({
    name: 'xl/workbook.xml',
    data: enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name).slice(0, 31)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>
</workbook>`),
  });

  files.push({
    name: 'xl/_rels/workbook.xml.rels',
    data: enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}
<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
  });

  files.push({ name: 'xl/styles.xml', data: enc(styles()) });
  sheets.forEach((s, i) => {
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: enc(sheetXml(s)) });
  });

  return zip(files);
}

export const STYLES = S;
