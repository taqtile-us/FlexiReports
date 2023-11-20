import generate from '../src/index';

generate({name: 'Name', assets: [{serialNumber: 1, jobs: [{title: 'title1', startDate: '2022-12-12', completionDate: '2022-12-14'}, {title: 'title2', startDate: '2022-12-12', completionDate: '2022-12-14'}]}, {serialNumber: 2, jobs: [{title: 'title1', startDate: '2022-12-12', completionDate: '2022-12-14'}, {title: 'title2', startDate: '2022-12-12', completionDate: '2022-12-14'}]}, {serialNumber: 3, jobs: [{title: 'title1', startDate: '2022-12-10', completionDate: '2022-12-14'}]}]}, 'with-diagramm.xlsx')
