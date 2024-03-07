import { writeDataToExcel } from './Templater';
import { parse } from './Parser';

export const generateExcelReport = async (dataToFill: any, filePath: string, reportPath: string, temporaryFolderPath: string) => {
  try {
    await writeDataToExcel(dataToFill, filePath, reportPath, temporaryFolderPath);

    return reportPath;
  } catch (e) {
    console.log(e, 'generateExcelReport error');
    return false;
  }
};

export {parse};