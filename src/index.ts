import {generateExcelReport} from "./generate-excel-report";
import {generateReportWithFormula, getCoordsOfFormulaCell} from "./Formuler";


const generate = async (dataToFill: {}, templatePath: string, reportPath: string) => {
    const report: any = await generateExcelReport(dataToFill, templatePath, reportPath);
    if (report) {
        const coords = await getCoordsOfFormulaCell(report);
        await generateReportWithFormula(report, coords);
    }
    return report;
}
export default generate;




