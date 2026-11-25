'use strict';

const spreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const officeRelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const xlsxMime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const xmlDeclaration = '<?xml version="1.0" encoding="UTF-8"?>';

function xmlText(value) {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

function workbookDate(value) {
  const year = Number(value.split('-')[0]);
  if (year < 1900 || year > 9999) return value;
  // Excel includes a fictitious leap day in 1900.
  const epoch = value < '1900-03-01' ? '1899-12-31' : '1899-12-30';
  return (Date.parse(value) - Date.parse(epoch)) / 86_400_000;
}

function worksheetXml(result) {
  const rows = [];
  const merged = [];
  const addRow = (values, styles = []) => {
    const row = rows.length + 1;
    const cells = values.map((value, index) => {
      const reference = `${index === 0 ? 'A' : 'B'}${row}`;
      const attributes = `r="${reference}" s="${styles[index] ?? 0}"`;
      if (typeof value === 'number') return `<c ${attributes}><v>${value}</v></c>`;
      return `<c ${attributes} t="inlineStr"><is><t>${xmlText(value)}</t></is></c>`;
    });
    rows.push(`<row r="${row}">${cells.join('')}</row>`);
  };
  const addWideRow = (value, style) => {
    addRow([value], [style]);
    merged.push(`<mergeCell ref="A${rows.length}:B${rows.length}"/>`);
  };

  addWideRow('국내 출장 여비 계산 결과', 1);
  addRow(['출장일', workbookDate(result.departure)], [0, 2]);
  addRow(['복귀일', workbookDate(result.returnDate)], [0, 2]);
  addRow(['출장 구분', result.tripType]);
  addRow(['출장일수', result.days]);
  addRow(['예상 지급액 (원)', result.total], [1, 4]);
  addRow([]);
  addRow(['항목', '금액 (원)'], [1, 1]);
  for (const item of result.items) addRow(item, [0, 3]);
  if (result.notes.length) {
    addRow([]);
    addWideRow('지급 조건 참고', 1);
    for (const note of result.notes) addWideRow(note, 0);
  }

  return `<worksheet xmlns="${spreadsheetNamespace}">
    <cols><col min="1" max="1" width="38" customWidth="1"/><col min="2" max="2" width="24" customWidth="1"/></cols>
    <sheetData>${rows.join('')}</sheetData>
    <mergeCells count="${merged.length}">${merged.join('')}</mergeCells>
  </worksheet>`;
}

async function downloadTripWorkbook(result) {
  const zip = new JSZip();
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    </Types>`,
    '_rels/.rels': `<Relationships xmlns="${packageRelationships}">
      <Relationship Id="rId1" Type="${officeRelationships}/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${spreadsheetNamespace}" xmlns:r="${officeRelationships}">
      <sheets><sheet name="출장 여비" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${packageRelationships}">
      <Relationship Id="rId1" Type="${officeRelationships}/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="${officeRelationships}/styles" Target="styles.xml"/>
    </Relationships>`,
    'xl/styles.xml': `<styleSheet xmlns="${spreadsheetNamespace}">
      <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
      <fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/></font></fonts>
      <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFE9F5EF"/><bgColor indexed="64"/></patternFill></fill></fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="5">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="3" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/>
      </cellXfs>
      <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
    </styleSheet>`,
    'xl/worksheets/sheet1.xml': worksheetXml(result),
  };
  for (const [path, xml] of Object.entries(parts)) zip.file(path, xmlDeclaration + xml);
  const blob = await zip.generateAsync({ type: 'blob', mimeType: xlsxMime });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `출장여비_${result.departure}_${result.returnDate}.xlsx`;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
