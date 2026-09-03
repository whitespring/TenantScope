import { trustedMicrosoftUrl } from './analysis.mjs';

const encoder = new TextEncoder();

const crcTable = Array.from({ length: 256 }, (_, start) => {
  let value = start;
  for (let bit = 0; bit < 8; bit += 1) value = (value & 1) ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function u16(value) {
  const bytes = new Uint8Array(2);
  new DataView(bytes.buffer).setUint16(0, value, true);
  return bytes;
}

function u32(value) {
  const bytes = new Uint8Array(4);
  new DataView(bytes.buffer).setUint32(0, value, true);
  return bytes;
}

function concat(...parts) {
  const result = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.length;
  }
  return result;
}

function zip(files) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const [name, content] of Object.entries(files)) {
    const nameBytes = encoder.encode(name);
    const data = encoder.encode(content);
    const checksum = crc32(data);
    const local = concat(
      u32(0x04034b50), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(checksum), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), nameBytes, data,
    );
    const central = concat(
      u32(0x02014b50), u16(20), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(checksum), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), u16(0),
      u16(0), u16(0), u32(0), u32(offset), nameBytes,
    );
    localParts.push(local);
    centralParts.push(central);
    offset += local.length;
  }

  const central = concat(...centralParts);
  return concat(
    ...localParts,
    central,
    u32(0x06054b50), u16(0), u16(0), u16(centralParts.length), u16(centralParts.length),
    u32(central.length), u32(offset), u16(0),
  );
}

function xmlEscape(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;',
  })[character]);
}

const pageWidth = 9360;

function runXml(text, { bold, color, size, style } = {}) {
  const properties = [style && `<w:rStyle w:val="${style}"/>`, bold && '<w:b/>', color && `<w:color w:val="${color}"/>`, size && `<w:sz w:val="${size}"/>`].filter(Boolean).join('');
  const content = String(text ?? '').split('\n').map((part, index) => `${index ? '<w:br/>' : ''}<w:t xml:space="preserve">${xmlEscape(part)}</w:t>`).join('');
  return `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}${content}</w:r>`;
}

function inlineXml(value, links, options = {}) {
  const parts = Array.isArray(value) ? value : [{ text: value }];
  return parts.map((part) => {
    if (typeof part === 'string') return runXml(part, options);
    if (!part.href) return runXml(part.text, { ...options, ...part });
    const href = trustedMicrosoftUrl(part.href);
    if (!href) return runXml(part.text, { ...options, ...part });
    const id = `rId${links.push(href) + 3}`;
    return `<w:hyperlink r:id="${id}">${runXml(part.text, { ...options, ...part, style: 'Hyperlink' })}</w:hyperlink>`;
  }).join('');
}

function paragraphXml(value, links, style = 'Normal', { align, keepNext, pageBreakBefore } = {}) {
  const properties = [
    `<w:pStyle w:val="${style}"/>`,
    align && `<w:jc w:val="${align}"/>`,
    keepNext && '<w:keepNext/>',
    pageBreakBefore && '<w:pageBreakBefore/>',
  ].filter(Boolean).join('');
  return `<w:p><w:pPr>${properties}</w:pPr>${inlineXml(value, links)}</w:p>`;
}

function tableProperties(width = pageWidth, border = 'DFE4E8', fill) {
  return `<w:tblPr><w:tblW w:w="${width}" w:type="dxa"/><w:tblInd w:w="120" w:type="dxa"/><w:tblLayout w:type="fixed"/>
    <w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((edge) => `<w:${edge} w:val="single" w:sz="6" w:color="${border}"/>`).join('')}</w:tblBorders>
    <w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>
    ${fill ? `<w:shd w:fill="${fill}"/>` : ''}</w:tblPr>`;
}

function cellXml(value, width, links, { fill, color, bold, style = 'TableText', align = 'left', paragraphs } = {}) {
  const content = paragraphs
    ? paragraphs.map((paragraph) => paragraphXml(paragraph.value, links, paragraph.style || style, { align: paragraph.align || align, keepNext: paragraph.keepNext })).join('')
    : paragraphXml(Array.isArray(value) ? value : [{ text: value, color, bold }], links, style, { align });
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="center"/>${fill ? `<w:shd w:fill="${fill}"/>` : ''}</w:tcPr>${content}</w:tc>`;
}

function simpleCellXml(value, width, links, { fill, color, bold, style = 'TableText', align = 'left' } = {}) {
  return cellXml(value, width, links, { fill, color, bold, style, align });
}

function widthsFor(count, supplied) {
  if (supplied?.length === count && supplied.reduce((sum, width) => sum + width, 0) === pageWidth) return supplied;
  const width = Math.floor(pageWidth / count);
  return Array.from({ length: count }, (_, index) => index === count - 1 ? pageWidth - width * (count - 1) : width);
}

function dataTableXml(block, links) {
  const widths = widthsFor(block.columns.length, block.widths);
  const alignments = block.alignments || block.columns.map((_, index) => index ? 'center' : 'left');
  const header = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${block.columns.map((column, index) => simpleCellXml(column, widths[index], links, { fill: '176B87', color: 'FFFFFF', bold: true, style: 'TableHeader', align: alignments[index] })).join('')}</w:tr>`;
  const rows = block.rows.map((row, rowIndex) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${row.map((value, index) => {
    const href = typeof value === 'string' && trustedMicrosoftUrl(value);
    const cell = href ? [{ text: 'Öffnen', href }] : value;
    return simpleCellXml(cell, widths[index], links, { fill: rowIndex % 2 ? 'F7F9FA' : 'FFFFFF', align: alignments[index] });
  }).join('')}</w:tr>`).join('');
  return `<w:tbl>${tableProperties()}<w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${header}${rows}</w:tbl>${paragraphXml('', links, 'Spacer')}`;
}

function metricStripXml(items, links) {
  const widths = widthsFor(items.length);
  const cells = items.map((item, index) => cellXml('', widths[index], links, {
    fill: index === 0 ? '102A3A' : 'F3F5F6',
    paragraphs: [
      { value: [{ text: item.label, color: index === 0 ? 'B8C8D0' : '66717C' }], style: 'MetricLabel', align: 'center', keepNext: true },
      { value: [{ text: item.value, color: index === 0 ? '7ED3CF' : '14202B', bold: true }], style: 'MetricValue', align: 'center' },
    ],
  })).join('');
  return `<w:tbl>${tableProperties(pageWidth, 'FFFFFF')}<w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr>${cells}</w:tr></w:tbl>${paragraphXml('', links, 'Spacer')}`;
}

function findingXml(block, links) {
  const tones = {
    high: ['FDE8E7', '9E2929'], medium: ['FFF2D8', '7A5813'], low: ['DFF3ED', '27644F'],
    error: ['EEE9E9', '6D5555'], info: ['DFF3ED', '27644F'], ok: ['DFF3ED', '27644F'],
  };
  const [fill, color] = tones[block.severity] || tones.info;
  const content = [
    { value: block.title, style: 'FindingTitle', keepNext: true },
    { value: [{ text: 'FESTSTELLUNG', color: '8A949B', bold: true }], style: 'FieldLabel', keepNext: true },
    { value: block.description, style: 'CompactBody' },
    ...(block.context ? [{ value: [{ text: 'EINORDNUNG', color: '8A949B', bold: true }], style: 'FieldLabel', keepNext: true }, { value: block.context, style: 'CompactBody' }] : []),
    ...(block.goodPractice ? [{ value: [{ text: 'GOOD PRACTICE', color: '8A949B', bold: true }], style: 'FieldLabel', keepNext: true }, { value: block.goodPractice, style: 'CompactBody' }] : []),
    { value: [{ text: 'EMPFOHLENE MASSNAHME', color: '8A949B', bold: true }], style: 'FieldLabel', keepNext: true },
    { value: [{ text: block.action, color: '0D5068', bold: true }], style: 'CompactBody' },
    ...(block.links?.length ? [{ value: block.links.flatMap((link, index) => [...(index ? [{ text: '  ·  ' }] : []), link]), style: 'CompactBody' }] : []),
  ];
  return `<w:tbl>${tableProperties(pageWidth, 'DFE4E8')}<w:tblGrid><w:gridCol w:w="1300"/><w:gridCol w:w="8060"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr>
    ${simpleCellXml(block.label, 1300, links, { fill, color, bold: true, style: 'Severity', align: 'center' })}
    ${cellXml('', 8060, links, { paragraphs: content })}
  </w:tr></w:tbl>${paragraphXml('', links, 'Spacer')}`;
}

function calloutXml(block, links) {
  const paragraphs = [
    { value: [{ text: block.label || 'GOOD PRACTICE', color: '176B87', bold: true }], style: 'FieldLabel', keepNext: true },
    ...(block.title ? [{ value: block.title, style: 'FindingTitle', keepNext: true }] : []),
    { value: block.text, style: 'CompactBody' },
    ...(block.links?.length ? [{ value: block.links.flatMap((link, index) => [...(index ? [{ text: '  ·  ' }] : []), link]), style: 'CompactBody' }] : []),
  ];
  return `<w:tbl>${tableProperties(pageWidth, '7ED3CF', 'F1FBFA')}<w:tblGrid><w:gridCol w:w="${pageWidth}"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr>${cellXml('', pageWidth, links, { fill: 'F1FBFA', paragraphs })}</w:tr></w:tbl>${paragraphXml('', links, 'Spacer')}`;
}

function metaXml(rows, links) {
  return `<w:tbl><w:tblPr><w:tblW w:w="${pageWidth}" w:type="dxa"/><w:tblInd w:w="120" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:insideH w:val="single" w:sz="4" w:color="DFE4E8"/></w:tblBorders><w:tblCellMar><w:top w:w="55" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="55" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1900"/><w:gridCol w:w="7460"/></w:tblGrid>${rows.map(([label, value]) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${simpleCellXml(label, 1900, links, { color: '66717C', bold: true, style: 'Metadata' })}${simpleCellXml(value, 7460, links, { color: '14202B', style: 'Metadata' })}</w:tr>`).join('')}</w:tbl>${paragraphXml('', links, 'Spacer')}`;
}

function blockXml(block, links) {
  if (block.type === 'metrics') return metricStripXml(block.items, links);
  if (block.type === 'finding') return findingXml(block, links);
  if (block.type === 'callout') return calloutXml(block, links);
  if (block.type === 'meta') return metaXml(block.rows, links);
  if (block.type === 'table') return dataTableXml(block, links);
  if (block.type === 'pageBreak') return paragraphXml('', links, 'Normal', { pageBreakBefore: true });
  return paragraphXml(block.href ? [{ text: block.text, href: block.href }] : block.text, links, block.style || 'Normal', { keepNext: /^Heading/.test(block.style || '') });
}

export function createDocx(title, blocks, { kicker = 'M365 GOVERNANCE ASSESSMENT', subtitle = '' } = {}) {
  const links = [];
  const body = [
    paragraphXml(kicker, links, 'Kicker', { keepNext: true }),
    paragraphXml(title, links, 'Title', { keepNext: true }),
    ...(subtitle ? [paragraphXml(subtitle, links, 'Subtitle')] : []),
    ...blocks.map((block) => blockXml(block, links)),
  ].join('');

  return zip({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
        <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
      </Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>`,
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
        ${links.map((href, index) => `<Relationship Id="rId${index + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlEscape(href)}" TargetMode="External"/>`).join('')}
      </Relationships>`,
    'word/document.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
        ${body}
        <w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/></w:sectPr>
      </w:body></w:document>`,
    'word/header1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:pStyle w:val="HeaderFooter"/><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs></w:pPr><w:r><w:t>TenantScope</w:t></w:r><w:r><w:tab/><w:t>M365 Governance Assessment</w:t></w:r></w:p></w:hdr>`,
    'word/footer1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:pStyle w:val="HeaderFooter"/><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">Seite </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
    'word/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:color w:val="14202B"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:color w:val="14202B"/></w:rPr><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Kicker"><w:name w:val="Kicker"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="176B87"/><w:spacing w:val="25"/></w:rPr><w:pPr><w:spacing w:before="0" w:after="100"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:sz w:val="56"/><w:color w:val="102A3A"/></w:rPr><w:pPr><w:spacing w:before="0" w:after="80"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="24"/><w:color w:val="66717C"/></w:rPr><w:pPr><w:spacing w:after="280"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:sz w:val="32"/><w:color w:val="176B87"/></w:rPr><w:pPr><w:keepNext/><w:spacing w:before="320" w:after="160"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:sz w:val="26"/><w:color w:val="176B87"/></w:rPr><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="1F4D78"/></w:rPr><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Metadata"><w:name w:val="Metadata"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="19"/></w:rPr><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="MetricLabel"><w:name w:val="Metric Label"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="16"/></w:rPr><w:pPr><w:spacing w:after="50"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="MetricValue"><w:name w:val="Metric Value"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="40"/></w:rPr><w:pPr><w:spacing w:after="0"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="FindingTitle"><w:name w:val="Finding Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="25"/><w:color w:val="14202B"/></w:rPr><w:pPr><w:keepNext/><w:spacing w:after="70"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="FieldLabel"><w:name w:val="Field Label"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="16"/><w:color w:val="8A949B"/></w:rPr><w:pPr><w:keepNext/><w:spacing w:before="80" w:after="25"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="CompactBody"><w:name w:val="Compact Body"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr><w:pPr><w:spacing w:after="70" w:line="250" w:lineRule="auto"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="TableHeader"><w:name w:val="Table Header"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="17"/><w:color w:val="FFFFFF"/></w:rPr><w:pPr><w:spacing w:after="0" w:line="230" w:lineRule="auto"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="Table Text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="17"/></w:rPr><w:pPr><w:spacing w:after="0" w:line="230" w:lineRule="auto"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Severity"><w:name w:val="Severity"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="17"/></w:rPr><w:pPr><w:spacing w:after="0"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Spacer"><w:name w:val="Spacer"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="6"/></w:rPr><w:pPr><w:spacing w:after="70"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="HeaderFooter"><w:name w:val="Header Footer"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="17"/><w:color w:val="7D878E"/></w:rPr><w:pPr><w:spacing w:after="0"/></w:pPr></w:style>
        <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="176B87"/><w:u w:val="single"/></w:rPr></w:style>
      </w:styles>`,
  });
}
