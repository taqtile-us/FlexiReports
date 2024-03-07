import Workbook from 'exceljs/index';
import excel from 'exceljs';
export async function readFile(path: string) {
    try {
        const workbook: Workbook = new excel.Workbook();
        await workbook.xlsx.readFile(path);
        return { workSheet: workbook.getWorksheet('template'), workbook };
    } catch (err) {
        console.error('Error reading the file:', err);
        return null;
    }
}