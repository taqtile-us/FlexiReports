import { writeDataToExcel } from './Templater';

export const generateExcelReport = async (dataToFill: any, filePath: string) => {
  try {
    const pathToReport = await writeDataToExcel(dataToFill, filePath);

    return pathToReport;
  } catch (e) {
    console.log(e, 'generateExcelReport error');
    return false;
  }
};
