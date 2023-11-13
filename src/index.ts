import {generateExcelReport} from "./generate-excel-report";


const generate = async (dataToFill: {}, filePath: string) => {
    const res: any = await generateExcelReport(dataToFill, filePath);
    return res;
}
export default generate;




