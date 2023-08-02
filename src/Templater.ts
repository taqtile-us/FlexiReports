import XlsxTemplate from 'xlsx-template';
import path from 'path';
import { readFile, writeFile } from 'fs/promises';

export const writeDataToExcel = async (dataToFill: any, templatePath: string) => {
  const filePath = 'uploads/report.xlsx';
  const data = await readFile(path.join(templatePath));

  const template = new XlsxTemplate(data);
  const sheetNumber = 1;

  template.substitute(sheetNumber, dataToFill);

  const binaryFileData = template.generate();
  await writeFile(filePath, binaryFileData, 'binary');
  return filePath;
};
