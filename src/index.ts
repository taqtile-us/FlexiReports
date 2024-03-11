import {generateExcelReport, parse, readFile} from "./generate-excel-report";

const generate = async (dataToFill: {}, filePath: string, reportPath: string, temporaryFolderPath: string) => {
    const res: any = await generateExcelReport(dataToFill, filePath, reportPath, temporaryFolderPath);
    return res;
}
export default generate;
export {parse, readFile}





