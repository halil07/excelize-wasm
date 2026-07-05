import { excelizeInit, File, ExcelizeAPI } from '../dist/excelize.cjs';

async function typeCheckCjs() {
  const Excelize: ExcelizeAPI = await excelizeInit();
  const file: File = Excelize.newFile();
  file.setCellValue('Sheet1', 'A1', 'cjs typed');
  const bytes: Uint8Array = file.save();
  file.close();

  // @ts-expect-error openFile should reject strings
  Excelize.openFile('not bytes');
}
