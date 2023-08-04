const XlsxTemplate = require('xlsx-template-ex');
import path from 'path';
import { writeFile } from 'fs/promises';

export const writeDataToExcel = async (dataToFill: any, templatePath: string, reportPath: string) => {
  const buffer: any = await XlsxTemplate.xlsxBuildByTemplate(dataToFill, path.join(templatePath));

  await writeFile(reportPath, buffer, 'binary');

  return reportPath;
};
