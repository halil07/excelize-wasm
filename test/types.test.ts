import { excelizeInit, File, ExcelizeAPI, Style, Chart, Table, DataValidation, PageLayoutOptions } from '../dist/excelize.mjs';

async function typeCheck() {
  const Excelize: ExcelizeAPI = await excelizeInit();

  // File lifecycle
  const file: File = Excelize.newFile();
  file.setCellValue('Sheet1', 'A1', 'typed');
  file.setCellValue('Sheet1', 'A2', 42);
  file.setCellValue('Sheet1', 'A3', true);
  file.setCellValue('Sheet1', 'A4', new Date());
  file.setCellValue('Sheet1', 'A5', null);

  const value: unknown = file.getCellValue('Sheet1', 'A1');
  const formula: string = file.getCellFormula('Sheet1', 'A2');

  file.setCellFormula('Sheet1', 'A6', '=SUM(A1:A5)');
  const calc: unknown = file.calcCellValue('Sheet1', 'A6');

  const bytes: Uint8Array = file.save();
  const reopened: File = Excelize.openFile(bytes);
  reopened.close();
  file.close();

  // Sheets
  const idx: number = file.newSheet('Data');
  file.setSheetName('Data', 'Renamed');
  file.setActiveSheet(idx);
  const active: number = file.getActiveSheetIndex();
  file.setSheetVisible('Renamed', false, true);
  file.deleteSheet('Renamed');

  // Rows / cols
  file.setSheetRow('Sheet1', 'A1', ['x', 1, true]);
  file.setSheetCol('Sheet1', 'B1', [1, 2, 3]);
  const rows: string[][] = file.getRows('Sheet1', { rawCellValue: true });
  const cols: string[][] = file.getCols('Sheet1');

  file.setColWidth('Sheet1', 'A', 'C', 20);
  const width: number = file.getColWidth('Sheet1', 'A');
  file.setColVisible('Sheet1', 'A', false);
  file.setRowHeight('Sheet1', 1, 30);
  const height: number = file.getRowHeight('Sheet1', 1);
  file.setRowVisible('Sheet1', 1, false);

  // Styles
  const style: Style = {
    font: { bold: true, size: 12 },
    fill: { type: 'pattern', color: ['#FFFF00'], pattern: 1 },
    alignment: { horizontal: 'center', wrapText: true },
    numFmt: 2,
  };
  const styleID: number = file.newStyle(style);
  file.setCellStyle('Sheet1', 'A1', 'A1', styleID);
  const appliedStyle: number = file.getCellStyle('Sheet1', 'A1');

  file.setConditionalFormat('Sheet1', 'A1:A2', [{
    type: 'cell',
    criteria: '>',
    value: '10',
    format: styleID,
  }]);

  // Merges
  file.mergeCell('Sheet1', 'B1', 'C2');
  file.unmergeCell('Sheet1', 'B1', 'C2');

  // Tables
  const table: Table = {
    name: 'Table1',
    range: 'A1:B2',
    styleName: 'TableStyleMedium2',
  };
  file.addTable('Sheet1', table);
  file.deleteTable('Table1');

  // Data validation
  const dv: DataValidation = {
    type: 'whole',
    operator: 'between',
    formula1: '1',
    formula2: '10',
    sqref: 'C1',
  };
  file.addDataValidation('Sheet1', dv);

  // Pictures
  const png = new Uint8Array([0x89, 0x50, 0x4E, 0x47]);
  file.addPictureFromBytes('Sheet1', 'D1', '.png', png, { scaleX: 0.5, scaleY: 0.5 });

  // Charts
  const chart: Chart = {
    type: 6,
    series: [{ name: 'Sheet1!$A$1', values: 'Sheet1!$B$1:$B$2' }],
    title: [{ text: 'Demo' }],
    xAxis: { title: [{ text: 'X' }] },
    yAxis: { title: [{ text: 'Y' }] },
  };
  file.addChart('Sheet1', 'F1', chart);
  file.addChartSheet('Chart1', chart);

  // Hyperlinks
  file.setCellHyperLink('Sheet1', 'A3', 'https://example.com', 'External');
  const link = file.getCellHyperLink('Sheet1', 'A3');

  // Layout / view / props
  const layout: PageLayoutOptions = { orientation: 'landscape' };
  file.setPageLayout('Sheet1', layout);
  file.setSheetView('Sheet1', 0, { zoomScale: 120 });
  file.setSheetProps('Sheet1', { outlineSymbols: true });

  // Document props
  file.setAppProps({ company: 'Example' });
  file.setDocProps({ title: 'Demo' });
  file.setCalcProps({ calcMode: 'auto' });

  // Raw bindings
  const rawFile = Excelize.raw.newFile(null);
  if (rawFile.ok) {
    const id: number = rawFile.data;
    Excelize.raw.setCellValue(id, 'Sheet1', 'A1', 'raw');
    const saved = Excelize.raw.save(id);
    if (saved.ok) {
      const rawBytes: Uint8Array = saved.data;
    }
    Excelize.raw.close(id);
  }

  // Type errors we want the compiler to catch (kept as comments):
  // @ts-expect-error openFile should reject strings
  Excelize.openFile('not bytes');
  // @ts-expect-error setCellStyle expects a number style ID
  file.setCellStyle('Sheet1', 'A1', 'A1', 'style');
}
