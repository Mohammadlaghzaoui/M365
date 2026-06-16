/**
 * Multi-sheet Excel (SpreadsheetML) export — no external libraries.
 * Produces a .xls workbook that opens natively in Excel, with one worksheet
 * per dataset. Used by the Migration Discovery module to fill the analysis
 * workbook from read-only Microsoft Graph data.
 */

export interface Sheet {
  name: string;
  columns: string[];
  rows: (string | number | boolean | null | undefined)[][];
}

function esc(v: string | number | boolean | null | undefined): string {
  return String(v ?? '').replace(/[<>&"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' }[c]!));
}

function cell(v: string | number | boolean | null | undefined): string {
  const isNum = typeof v === 'number' && Number.isFinite(v);
  const type = isNum ? 'Number' : 'String';
  const val = isNum ? v : esc(v);
  return `<Cell><Data ss:Type="${type}">${val}</Data></Cell>`;
}

export function buildWorkbook(sheets: Sheet[]): string {
  const worksheets = sheets.map((s) => {
    const header = `<Row>${s.columns.map((c) => `<Cell ss:StyleID="hdr"><Data ss:Type="String">${esc(c)}</Data></Cell>`).join('')}</Row>`;
    const body = s.rows.map((r) => `<Row>${r.map(cell).join('')}</Row>`).join('');
    const name = esc(s.name).slice(0, 31).replace(/[\\/?*[\]:]/g, ' ');
    return `<Worksheet ss:Name="${name}"><Table>${header}${body}</Table></Worksheet>`;
  }).join('');

  return `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Styles>
  <Style ss:ID="hdr"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#2563EB" ss:Pattern="Solid"/></Style>
 </Styles>
 ${worksheets}
</Workbook>`;
}

export function downloadWorkbook(sheets: Sheet[], filename: string): void {
  const xml = buildWorkbook(sheets);
  const blob = new Blob([xml], { type: 'application/vnd.ms-excel' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename.endsWith('.xls') ? filename : `${filename}.xls`;
  a.click();
  URL.revokeObjectURL(a.href);
}
