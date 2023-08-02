import {generateExcelReport} from "./generate-excel-report";
import {generateReportWithFormula, getCoordsOfFormulaCell} from "./Formuler";


const generate = async (dataToFill: {}, filePath: string) => {
    const res: any = await generateExcelReport(dataToFill, filePath);
    console.log(res, 'res');
    if (res) {
        const coords = await getCoordsOfFormulaCell(res);
        console.log(coords, 'coords');
        await generateReportWithFormula(res, coords);
    }
    return res;
}
export default generate;




