import generate from '../src/index'
import { promises as fsPromises } from 'fs';
const templatePath = 'templates/with-diagram.xlsx';
const reportPath = 'report.xlsx';
const temporaryFolder = 'temporary';
import fs from 'fs';


const start = async () => {
  fs.rm(temporaryFolder, { recursive: true, force: true }, async (err) => {
    if (err) {
      throw err;
    }

    console.log(`${temporaryFolder} is deleted!`);
    fs.unlink('report.xlsx', async () => {
      console.log('report file is deleted')
      const jsonData = await fsPromises.readFile('assetClass.json', 'utf8');
      const parsedData = JSON.parse(jsonData);
      generate(parsedData, templatePath, reportPath, temporaryFolder);
    });
  });

};
start();
// generate({name: 'Name', startDatePeriod: '2022-12-12', assets: [{serialNumber: 1, jobs: [{title: 'title1', startDate: '2022-12-12', completionDate: '2022-12-14'}, {title: 'title2', startDate: '2022-12-12', completionDate: '2022-12-14'}]}, {serialNumber: 2, jobs: [{title: 'title1', startDate: '2022-12-12', completionDate: '2022-12-14'}, {title: 'title2', startDate: '2022-12-12', completionDate: '2022-12-14'}]}, {serialNumber: 3, jobs: [{title: 'title1', startDate: '2022-12-10', completionDate: '2022-12-14'}]}]}, 'with-diagramm.xlsx')
