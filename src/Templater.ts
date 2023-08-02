import XlsxTemplate from 'xlsx-template-ex';
import path from 'path';
import { writeFile } from 'fs/promises';

export const writeDataToExcel = async (dataToFill: any, templatePath: string) => {
  const filePath = 'uploads/report.xlsx';
  const buffer: any = await XlsxTemplate.xlsxBuildByTemplate(dataToFill, path.join(templatePath));

  await writeFile(filePath, buffer, 'binary');

  return filePath;
};
