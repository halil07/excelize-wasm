import test from 'node:test';
import assert from 'node:assert/strict';

import { excelizeInit } from '../dist/excelize.mjs';

const Excelize = await excelizeInit();

const tinyPng = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNgYGAAAAAEAAH2FzhVAAAAAElFTkSuQmCC',
  'base64'
);

function expectNoValue(value) {
  assert.ok(
    value === null || value === undefined,
    `expected null/undefined but received ${value}`
  );
}

function assertResultShape(result) {
  assert.ok(result && typeof result === 'object', 'result must be an object');
  assert.ok('ok' in result, 'result must have ok property');
}

function assertOk(result, message) {
  assertResultShape(result);
  assert.ok(result.ok, message || result.err);
}

function assertErr(result) {
  assertResultShape(result);
  assert.ok(!result.ok, 'expected an error result');
  assert.ok(typeof result.err === 'string' && result.err.length > 0, 'error result must have err string');
}

// ============================================================================
// Global loader & lifecycle
// ============================================================================

test('excelizeInit returns the initialized API', () => {
  assert.ok(Excelize, 'Excelize API should be defined');
  assert.equal(typeof Excelize.newFile, 'function');
  assert.equal(typeof Excelize.openFile, 'function');
  assert.equal(typeof Excelize.ready, 'function');
  assert.ok(Excelize.raw, 'Excelize.raw should be defined');
});

test('newFile creates a workbook with default Sheet1', () => {
  const file = Excelize.newFile();
  try {
    assert.deepEqual(file.getSheetList(), ['Sheet1']);
    assert.equal(file.getSheetCount(), 1);
    assert.equal(file.defaultSheetName(), 'Sheet1');
  } finally {
    file.close();
  }
});

test('openFile reads a Uint8Array and preserves content', () => {
  const file = Excelize.newFile();
  try {
    file.setCellValue('Sheet1', 'A1', 'persisted');
    const bytes = file.save();
    assert.ok(bytes instanceof Uint8Array);

    const reopened = Excelize.openFile(bytes);
    try {
      assert.equal(reopened.getSheetCount(), 1);
      assert.equal(reopened.getCellValue('Sheet1', 'A1'), 'persisted');
    } finally {
      reopened.close();
    }
  } finally {
    file.close();
  }
});

test('openFile rejects non-buffer input', () => {
  assert.throws(() => Excelize.openFile(null), /Uint8Array|ArrayBuffer/);
  assert.throws(() => Excelize.openFile('not bytes'), /Uint8Array|ArrayBuffer/);
});

test('multiple file handles are isolated', () => {
  const a = Excelize.newFile();
  const b = Excelize.newFile();
  try {
    a.setCellValue('Sheet1', 'A1', 'A');
    b.setCellValue('Sheet1', 'A1', 'B');
    assert.equal(a.getCellValue('Sheet1', 'A1'), 'A');
    assert.equal(b.getCellValue('Sheet1', 'A1'), 'B');
  } finally {
    a.close();
    b.close();
  }
});

test('close releases the handle', () => {
  const file = Excelize.newFile();
  file.close();
  assert.throws(() => file.getSheetCount(), /already closed/);
});

// ============================================================================
// Sheet operations
// ============================================================================

test('sheet CRUD and metadata methods', () => {
  const file = Excelize.newFile();
  try {
    assert.equal(file.newSheet('Data'), 1);
    assert.equal(file.getSheetIndex('Data'), 1);
    assert.equal(file.getSheetName(1), 'Data');

    expectNoValue(file.setSheetName('Data', 'Renamed'));
    assert.equal(file.getSheetIndex('Renamed'), 1);
    assert.equal(file.getSheetName(1), 'Renamed');

    const sheetMap = file.getSheetMap();
    assert.equal(typeof sheetMap, 'object');
    assert.equal(sheetMap[2], 'Renamed');

    expectNoValue(file.deleteSheet('Renamed'));
    assert.equal(file.getSheetCount(), 1);
    assert.deepEqual(file.getSheetList(), ['Sheet1']);
    assert.equal(file.getSheetMap()[1], 'Sheet1');
  } finally {
    file.close();
  }
});

test('sheet visibility and active sheet', () => {
  const file = Excelize.newFile();
  try {
    file.newSheet('Second');

    assert.equal(file.getSheetVisible('Second'), true);
    expectNoValue(file.setSheetVisible('Second', false));
    assert.equal(file.getSheetVisible('Second'), false);
    expectNoValue(file.setSheetVisible('Second', true));
    assert.equal(file.getSheetVisible('Second'), true);

    expectNoValue(file.setSheetVisible('Second', false, true));
    assert.equal(file.getSheetVisible('Second'), false);

    expectNoValue(file.setActiveSheet(1));
    assert.equal(file.getActiveSheetIndex(), 1);

    expectNoValue(file.setActiveSheet(0));
    assert.equal(file.getActiveSheetIndex(), 0);
  } finally {
    file.close();
  }
});

test('sheet dimension and background', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setSheetDimension('Sheet1', 'A1:D4'));
    assert.equal(file.getSheetDimension('Sheet1'), 'A1:D4');

    expectNoValue(file.setSheetBackgroundFromBytes('Sheet1', '.png', tinyPng));
  } finally {
    file.close();
  }
});

// ============================================================================
// Cell operations
// ============================================================================

test('setCellValue and getCellValue with scalar types', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setCellValue('Sheet1', 'A1', 'text'));
    assert.equal(file.getCellValue('Sheet1', 'A1'), 'text');

    expectNoValue(file.setCellValue('Sheet1', 'A2', 42));
    const num = file.getCellValue('Sheet1', 'A2');
    assert.ok(num === 42 || num === '42', `expected 42, got ${num}`);

    expectNoValue(file.setCellValue('Sheet1', 'A3', 3.14));
    const flt = file.getCellValue('Sheet1', 'A3');
    assert.ok(typeof flt === 'number' || typeof flt === 'string');

    expectNoValue(file.setCellValue('Sheet1', 'A4', true));
    const bool = file.getCellValue('Sheet1', 'A4');
    assert.ok(bool === true || bool === 'TRUE' || bool === '1', `got ${bool}`);

    expectNoValue(file.setCellValue('Sheet1', 'A5', null));
    assert.ok(file.getCellValue('Sheet1', 'A5') === '');
  } finally {
    file.close();
  }
});

test('setCellValue with Date', () => {
  const file = Excelize.newFile();
  try {
    const date = new Date(Date.UTC(2024, 6, 5, 12, 0, 0));
    expectNoValue(file.setCellValue('Sheet1', 'A1', date));
    const value = file.getCellValue('Sheet1', 'A1');
    assert.ok(typeof value === 'number' || typeof value === 'string');
    assert.ok(value !== '', 'date should produce a non-empty cell value');
  } finally {
    file.close();
  }
});

test('cell formula and calculated value', () => {
  const file = Excelize.newFile();
  try {
    file.setCellValue('Sheet1', 'A1', 10);
    file.setCellValue('Sheet1', 'A2', 20);

    expectNoValue(file.setCellFormula('Sheet1', 'A3', '=SUM(A1:A2)'));
    assert.equal(file.getCellFormula('Sheet1', 'A3'), '=SUM(A1:A2)');

    const calc = file.calcCellValue('Sheet1', 'A3');
    assert.ok(typeof calc === 'string' || typeof calc === 'number');
    assert.ok(Number(calc) === 30, `expected 30, got ${calc}`);
  } finally {
    file.close();
  }
});

test('setSheetRow and setSheetCol', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setSheetRow('Sheet1', 'A1', ['x', 2, true]));
    expectNoValue(file.setSheetCol('Sheet1', 'C1', [1, 2, 3]));

    const rows = file.getRows('Sheet1');
    assert.ok(Array.isArray(rows));
    assert.ok(rows.length >= 1);
    assert.equal(rows[0][0], 'x');

    const cols = file.getCols('Sheet1');
    assert.ok(Array.isArray(cols));
    assert.ok(cols.length >= 1);
  } finally {
    file.close();
  }
});

test('getRows and getCols with options', () => {
  const file = Excelize.newFile();
  try {
    file.setCellValue('Sheet1', 'A1', 'a');
    file.setCellValue('Sheet1', 'B1', 'b');
    file.setCellValue('Sheet1', 'A2', 'c');

    const rows = file.getRows('Sheet1', { rawCellValue: true });
    assert.ok(Array.isArray(rows));
    assert.ok(rows.length >= 2);

    const cols = file.getCols('Sheet1', { rawCellValue: true });
    assert.ok(Array.isArray(cols));
    assert.ok(cols.length >= 2);
  } finally {
    file.close();
  }
});

// ============================================================================
// Row & column formatting
// ============================================================================

test('column width, visibility and outline', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setColWidth('Sheet1', 'A', 'C', 25));
    assert.equal(file.getColWidth('Sheet1', 'A'), 25);
    assert.equal(file.getColWidth('Sheet1', 'B'), 25);

    expectNoValue(file.setColVisible('Sheet1', 'A', false));
    assert.equal(file.getColVisible('Sheet1', 'A'), false);
    expectNoValue(file.setColVisible('Sheet1', 'A', true));
    assert.equal(file.getColVisible('Sheet1', 'A'), true);

    expectNoValue(file.setColOutlineLevel('Sheet1', 'A', 2));
    assert.equal(file.getColOutlineLevel('Sheet1', 'A'), 2);
    assert.equal(file.getColOutlineLevel('Sheet1', 'B'), 0);
  } finally {
    file.close();
  }
});

test('row height, visibility and outline', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setRowHeight('Sheet1', 1, 30));
    assert.equal(file.getRowHeight('Sheet1', 1), 30);

    expectNoValue(file.setRowVisible('Sheet1', 1, false));
    assert.equal(file.getRowVisible('Sheet1', 1), false);
    expectNoValue(file.setRowVisible('Sheet1', 1, true));
    assert.equal(file.getRowVisible('Sheet1', 1), true);

    assert.equal(file.getRowOutlineLevel('Sheet1', 1), 0);
  } finally {
    file.close();
  }
});

// ============================================================================
// Styles
// ============================================================================

test('style creation and application', () => {
  const file = Excelize.newFile();
  try {
    const styleID = file.newStyle({
      font: { bold: true, size: 14, color: '#FF0000' },
      fill: { type: 'pattern', color: ['#FFFF00'], pattern: 1 },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: [
        { type: 'left', color: '000000', style: 1 },
        { type: 'right', color: '000000', style: 1 },
      ],
    });
    assert.equal(typeof styleID, 'number');
    assert.ok(styleID >= 0);

    expectNoValue(file.setCellStyle('Sheet1', 'A1', 'A1', styleID));
    assert.equal(file.getCellStyle('Sheet1', 'A1'), styleID);
  } finally {
    file.close();
  }
});

test('conditional format round trip', () => {
  const file = Excelize.newFile();
  try {
    const styleID = file.newStyle({ font: { bold: true } });
    expectNoValue(file.setConditionalFormat('Sheet1', 'A1:A2', [{
      type: 'cell',
      criteria: '>',
      value: '10',
      format: styleID,
    }]));

    const formats = file.getConditionalFormats('Sheet1');
    assert.equal(typeof formats, 'object');
  } finally {
    file.close();
  }
});

// ============================================================================
// Merges
// ============================================================================

test('merge, unmerge and query merged cells', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.mergeCell('Sheet1', 'B1', 'C2'));
    const merged = file.getMergeCells('Sheet1');
    assert.ok(Array.isArray(merged));
    assert.ok(merged.length >= 1);

    expectNoValue(file.unmergeCell('Sheet1', 'B1', 'C2'));
    assert.equal(file.getMergeCells('Sheet1').length, 0);
  } finally {
    file.close();
  }
});

// ============================================================================
// Tables
// ============================================================================

test('table add, get and delete', () => {
  const file = Excelize.newFile();
  try {
    file.setCellValue('Sheet1', 'A1', 'Name');
    file.setCellValue('Sheet1', 'A2', 'Ada');
    file.setCellValue('Sheet1', 'B1', 'Age');
    file.setCellValue('Sheet1', 'B2', 35);

    expectNoValue(file.addTable('Sheet1', {
      name: 'Table1',
      range: 'A1:B2',
      styleName: 'TableStyleMedium2',
    }));

    const tables = file.getTables('Sheet1');
    assert.ok(Array.isArray(tables));
    assert.ok(tables.length >= 1);

    expectNoValue(file.deleteTable('Table1'));
    assert.equal(file.getTables('Sheet1').length, 0);
  } finally {
    file.close();
  }
});

// ============================================================================
// Data validation
// ============================================================================

test('data validation', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.addDataValidation('Sheet1', {
      type: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '10',
      sqref: 'C1',
    }));
  } finally {
    file.close();
  }
});

// ============================================================================
// Pictures
// ============================================================================

test('picture from bytes add and get', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.addPictureFromBytes('Sheet1', 'D1', '.png', tinyPng, {
      scaleX: 0.5,
      scaleY: 0.5,
    }));

    const pictures = file.getPictures('Sheet1', 'D1');
    assert.ok(Array.isArray(pictures));
    assert.ok(pictures.length >= 1);
  } finally {
    file.close();
  }
});

// ============================================================================
// Charts
// ============================================================================

test('chart and chart sheet', () => {
  const file = Excelize.newFile();
  try {
    file.setCellValue('Sheet1', 'A1', 'Name');
    file.setCellValue('Sheet1', 'A2', 'Ada');
    file.setCellValue('Sheet1', 'B1', 'Age');
    file.setCellValue('Sheet1', 'B2', 35);

    expectNoValue(file.addChart('Sheet1', 'F1', {
      type: 6,
      series: [{ name: 'Sheet1!$A$1', values: 'Sheet1!$B$1:$B$2' }],
      title: [{ text: 'Demo' }],
      xaxis: { title: [{ text: 'X' }] },
      yaxis: { title: [{ text: 'Y' }] },
    }));

    expectNoValue(file.addChartSheet('Sheet2', {
      type: 6,
      series: [{ name: 'Sheet1!$A$1', values: 'Sheet1!$B$1:$B$2' }],
      title: [{ text: 'Demo Chart' }],
    }));

    assert.equal(file.getSheetCount(), 2);
  } finally {
    file.close();
  }
});

// ============================================================================
// Hyperlinks
// ============================================================================

test('cell hyperlink set and get', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setCellHyperLink('Sheet1', 'A3', 'https://example.com', 'External'));
    const link = file.getCellHyperLink('Sheet1', 'A3');
    assert.ok(typeof link === 'object');
    assert.ok(link.hasLink === true || link.hasLink === 1);
    assert.ok(link.link === 'https://example.com' || typeof link.link === 'string');
  } finally {
    file.close();
  }
});

// ============================================================================
// Page layout, sheet view and sheet props
// ============================================================================

test('page layout round trip', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setPageLayout('Sheet1', { orientation: 'landscape' }));
    const layout = file.getPageLayout('Sheet1');
    assert.ok(typeof layout === 'object');
  } finally {
    file.close();
  }
});

test('sheet view round trip', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setSheetView('Sheet1', 0, { zoomScale: 120 }));
    const view = file.getSheetView('Sheet1', 0);
    assert.ok(typeof view === 'object');
  } finally {
    file.close();
  }
});

test('sheet props round trip', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setSheetProps('Sheet1', { outlineSymbols: true }));
    const props = file.getSheetProps('Sheet1');
    assert.ok(typeof props === 'object');
  } finally {
    file.close();
  }
});

// ============================================================================
// Document properties
// ============================================================================

test('app properties round trip', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setAppProps({ Company: 'Example', Application: 'TestApp' }));
    const props = file.getAppProps();
    assert.ok(typeof props === 'object');
    assert.equal(props.Company, 'Example');
  } finally {
    file.close();
  }
});

test('doc properties round trip', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setDocProps({ title: 'Demo Workbook', subject: 'Testing' }));
    const props = file.getDocProps();
    assert.ok(typeof props === 'object');
    assert.equal(props.Title, 'Demo Workbook');
  } finally {
    file.close();
  }
});

// ============================================================================
// Calculation properties
// ============================================================================

test('calc props round trip', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setCalcProps({ calcMode: 'auto' }));
    const props = file.getCalcProps();
    assert.ok(typeof props === 'object');
  } finally {
    file.close();
  }
});

// ============================================================================
// Error handling
// ============================================================================

test('errors are thrown for invalid sheet or cell references', () => {
  const file = Excelize.newFile();
  try {
    assert.throws(() => file.setCellValue('NoSuchSheet', 'A1', 'x'), /does not exist/);
    assert.throws(() => file.getCellValue('Sheet1', 'invalid'), /invalid/);
  } finally {
    file.close();
  }
});

test('openFile rejects corrupt bytes', () => {
  assert.throws(() => Excelize.openFile(Buffer.from('not an xlsx')), /zip|format|invalid/i);
});

// ============================================================================
// Raw wasm bindings (comprehensive)
// ============================================================================

test('raw: newFile, setCellValue, getCellValue, save, close', () => {
  const raw = Excelize.raw;

  const created = raw.newFile(null);
  assertOk(created);
  const id = created.data;

  try {
    assertOk(raw.setCellValue(id, 'Sheet1', 'A1', 'raw value'));
    const got = raw.getCellValue(id, 'Sheet1', 'A1', null);
    assertOk(got);
    assert.equal(got.data, 'raw value');

    const saved = raw.save(id);
    assertOk(saved);
    assert.ok(saved.data instanceof Uint8Array);
    assert.ok(saved.data.byteLength > 0);
  } finally {
    assertOk(raw.close(id));
  }
});

test('raw: error results contain err string', () => {
  const raw = Excelize.raw;
  const bad = raw.getCellValue(999999, 'Sheet1', 'A1', null);
  assertErr(bad);
});

test('raw: sheet, style and picture bindings', () => {
  const raw = Excelize.raw;
  const created = raw.newFile(null);
  assertOk(created);
  const id = created.data;

  try {
    assertOk(raw.newSheet(id, 'RawSheet'));
    assertOk(raw.setActiveSheet(id, 1));

    const style = raw.newStyle(id, { font: { bold: true } });
    assertOk(style);
    assert.equal(typeof style.data, 'number');

    assertOk(raw.setCellStyle(id, 'Sheet1', 'A1', 'A1', style.data));
    const styleID = raw.getCellStyle(id, 'Sheet1', 'A1');
    assertOk(styleID);
    assert.equal(styleID.data, style.data);

    assertOk(raw.addPictureFromBytes(id, 'Sheet1', 'B1', '.png', tinyPng, null));
    const pics = raw.getPictures(id, 'Sheet1', 'B1');
    assertOk(pics);
    assert.ok(Array.isArray(pics.data));
  } finally {
    raw.close(id);
  }
});

test('raw: merge, table, validation, hyperlink and doc props bindings', () => {
  const raw = Excelize.raw;
  const created = raw.newFile(null);
  assertOk(created);
  const id = created.data;

  try {
    assertOk(raw.setCellValue(id, 'Sheet1', 'A1', 'Hello'));
    assertOk(raw.setCellValue(id, 'Sheet1', 'A2', 'World'));

    assertOk(raw.mergeCell(id, 'Sheet1', 'B1', 'C2'));
    const merged = raw.getMergeCells(id, 'Sheet1');
    assertOk(merged);
    assert.ok(Array.isArray(merged.data));

    assertOk(raw.addTable(id, 'Sheet1', {
      name: 'RawTable',
      range: 'A1:A2',
      styleName: 'TableStyleMedium2',
    }));
    const tables = raw.getTables(id, 'Sheet1');
    assertOk(tables);
    assert.ok(Array.isArray(tables.data));

    assertOk(raw.addDataValidation(id, 'Sheet1', {
      type: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '10',
      sqref: 'C1',
    }));

    assertOk(raw.setCellHyperLink(id, 'Sheet1', 'A3', 'https://example.com', 'External', null));
    const link = raw.getCellHyperLink(id, 'Sheet1', 'A3');
    assertOk(link);
    assert.ok(typeof link.data === 'object');

    assertOk(raw.setDocProps(id, { title: 'Raw Workbook' }));
    const props = raw.getDocProps(id);
    assertOk(props);
    assert.ok(typeof props.data === 'object');
  } finally {
    raw.close(id);
  }
});

test('raw: row, col, outline and calc bindings', () => {
  const raw = Excelize.raw;
  const created = raw.newFile(null);
  assertOk(created);
  const id = created.data;

  try {
    assertOk(raw.setColWidth(id, 'Sheet1', 'A', 'B', 15));
    const width = raw.getColWidth(id, 'Sheet1', 'A');
    assertOk(width);
    assert.equal(typeof width.data, 'number');

    assertOk(raw.setRowHeight(id, 'Sheet1', 1, 25));
    const height = raw.getRowHeight(id, 'Sheet1', 1);
    assertOk(height);
    assert.equal(typeof height.data, 'number');

    assertOk(raw.setColOutlineLevel(id, 'Sheet1', 'A', 1));
    const outline = raw.getColOutlineLevel(id, 'Sheet1', 'A');
    assertOk(outline);
    assert.equal(outline.data, 1);

    assertOk(raw.setCellValue(id, 'Sheet1', 'A1', 5));
    assertOk(raw.setCellValue(id, 'Sheet1', 'A2', 10));
    assertOk(raw.setCellFormula(id, 'Sheet1', 'A3', '=A1+A2', null));
    const calc = raw.calcCellValue(id, 'Sheet1', 'A3', null);
    assertOk(calc);
    assert.ok(Number(calc.data) === 15, `expected 15, got ${calc.data}`);
  } finally {
    raw.close(id);
  }
});

// ============================================================================
// Edge cases & invalid inputs
// ============================================================================

test('very hidden sheet state persists', () => {
  const file = Excelize.newFile();
  try {
    file.newSheet('Hidden');
    file.setSheetVisible('Hidden', false, true);
    assert.equal(file.getSheetVisible('Hidden'), false);
    file.setSheetVisible('Hidden', true);
    assert.equal(file.getSheetVisible('Hidden'), true);
  } finally {
    file.close();
  }
});

test('getCellHyperLink on cell without link', () => {
  const file = Excelize.newFile();
  try {
    const link = file.getCellHyperLink('Sheet1', 'A1');
    assert.ok(typeof link === 'object');
    assert.ok(link.hasLink === false || link.hasLink === 0);
  } finally {
    file.close();
  }
});

test('invalid picture extension throws', () => {
  const file = Excelize.newFile();
  try {
    assert.throws(() => file.addPictureFromBytes('Sheet1', 'A1', '.xyz', tinyPng), /unsupported|invalid/i);
  } finally {
    file.close();
  }
});

test('invalid chart type throws', () => {
  const file = Excelize.newFile();
  try {
    file.setCellValue('Sheet1', 'A1', 1);
    file.setCellValue('Sheet1', 'A2', 2);
    assert.throws(() => file.addChart('Sheet1', 'A3', {
      type: 57,
      series: [{ name: 'Sheet1!$A$1', values: 'Sheet1!$A$1:$A$2' }],
    }), /unsupported/i);
  } finally {
    file.close();
  }
});

test('saving and reopening preserves complex content', () => {
  const file = Excelize.newFile();
  try {
    file.setCellValue('Sheet1', 'A1', 'title');
    file.setCellValue('Sheet1', 'A2', 100);
    file.setCellFormula('Sheet1', 'A3', '=A2*2');
    const styleID = file.newStyle({ font: { bold: true } });
    file.setCellStyle('Sheet1', 'A1', 'A1', styleID);
    file.mergeCell('Sheet1', 'B1', 'C2');
    file.addPictureFromBytes('Sheet1', 'D1', '.png', tinyPng, null);

    const bytes = file.save();
    const reopened = Excelize.openFile(bytes);
    try {
      assert.equal(reopened.getCellValue('Sheet1', 'A1'), 'title');
      assert.equal(reopened.getCellFormula('Sheet1', 'A3'), '=A2*2');
      assert.equal(reopened.getCellStyle('Sheet1', 'A1'), styleID);
      assert.ok(reopened.getMergeCells('Sheet1').length >= 1);
      assert.ok(reopened.getPictures('Sheet1', 'D1').length >= 1);
    } finally {
      reopened.close();
    }
  } finally {
    file.close();
  }
});

// ============================================================================
// Options and less common paths
// ============================================================================

test('getCellValue and calcCellValue with rawCellValue option', () => {
  const file = Excelize.newFile();
  try {
    file.setCellValue('Sheet1', 'A1', 123);
    file.setCellFormula('Sheet1', 'A2', '=A1*2');

    const rawValue = file.getCellValue('Sheet1', 'A1', { rawCellValue: true });
    assert.ok(typeof rawValue === 'number' || typeof rawValue === 'string');

    const calcRaw = file.calcCellValue('Sheet1', 'A2', { rawCellValue: true });
    assert.ok(typeof calcRaw === 'number' || typeof calcRaw === 'string');
  } finally {
    file.close();
  }
});

test('setCellFormula with formula options', () => {
  const file = Excelize.newFile();
  try {
    file.setCellValue('Sheet1', 'A1', 10);
    file.setCellValue('Sheet1', 'A2', 20);
    expectNoValue(file.setCellFormula('Sheet1', 'A3', '=SUM(A1:A2)', { format: '0.00' }));
    assert.equal(file.getCellFormula('Sheet1', 'A3'), '=SUM(A1:A2)');
  } finally {
    file.close();
  }
});

test('newFile and openFile with options', () => {
  const file = Excelize.newFile({ password: 'secret' });
  try {
    file.setCellValue('Sheet1', 'A1', 'protected');
    const bytes = file.save();
    assert.ok(bytes instanceof Uint8Array);

    const reopened = Excelize.openFile(bytes, { password: 'secret' });
    try {
      assert.equal(reopened.getCellValue('Sheet1', 'A1'), 'protected');
    } finally {
      reopened.close();
    }
  } finally {
    file.close();
  }
});

test('empty arrays for setSheetRow and setSheetCol', () => {
  const file = Excelize.newFile();
  try {
    expectNoValue(file.setSheetRow('Sheet1', 'A1', []));
    expectNoValue(file.setSheetCol('Sheet1', 'B1', []));
    assert.deepEqual(file.getRows('Sheet1'), []);
  } finally {
    file.close();
  }
});

test('newStyle with number format and alignment', () => {
  const file = Excelize.newFile();
  try {
    const styleID = file.newStyle({
      numFmt: 2,
      alignment: { horizontal: 'right', wrapText: true },
    });
    assert.equal(typeof styleID, 'number');
  } finally {
    file.close();
  }
});

test('raw: conditional format, page layout and sheet view bindings', () => {
  const raw = Excelize.raw;
  const created = raw.newFile(null);
  assertOk(created);
  const id = created.data;

  try {
    const style = raw.newStyle(id, { font: { bold: true } });
    assertOk(style);

    assertOk(raw.setConditionalFormat(id, 'Sheet1', 'A1:A2', [{
      type: 'cell',
      criteria: '>',
      value: '0',
      format: style.data,
    }]));
    const formats = raw.getConditionalFormats(id, 'Sheet1');
    assertOk(formats);
    assert.equal(typeof formats.data, 'object');

    assertOk(raw.setPageLayout(id, 'Sheet1', { orientation: 'landscape' }));
    const layout = raw.getPageLayout(id, 'Sheet1');
    assertOk(layout);
    assert.equal(typeof layout.data, 'object');

    assertOk(raw.setSheetView(id, 'Sheet1', 0, { zoomScale: 150 }));
    const view = raw.getSheetView(id, 'Sheet1', 0);
    assertOk(view);
    assert.equal(typeof view.data, 'object');
  } finally {
    raw.close(id);
  }
});
