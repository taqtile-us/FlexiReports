npm i xlsx-constructor

```
import generate from 'xlsx-constructor';

const data = {name: 'Name', startDatePeriod: '2022-12-12', jobs: [{title: 'title1', startDate: '2022-12-10', completionDate: '2022-12-14'}, {title: 'title2', startDate: '2022-12-13', completionDate: '2022-12-15'}],  assets: [{serialNumber: 1, jobs: [{title: 'title1', startDate: '2022-12-12', completionDate: '2022-12-14'}, {title: 'title2', startDate: '2022-12-12', completionDate: '2022-12-14'}]}, {serialNumber: 2, jobs: [{title: 'title1', startDate: '2022-12-12', completionDate: '2022-12-14'}, {title: 'title2', startDate: '2022-12-12', completionDate: '2022-12-14'}]}, {serialNumber: 3, jobs: [{title: 'title1', startDate: '2022-12-10', completionDate: '2022-12-14'}]}]}
const templatePath = 'templates/with-diagramm.xlsx';
const reportPath = 'report.xlsx';
const temporaryFolder = 'temporary';
generate(data, templatePath, reportPath, temporaryFolder)
```
