const readXlsxFile = require('read-excel-file/node');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;

String.prototype.replaceAll = function (match, replace) {
  return this.replace(new RegExp(match, 'g'), () => replace);
};

const PATH_CSV = './data.csv';
const PATH_DATA = './data_20210614.xlsx';

const csvWriter = createCsvWriter({
  path: PATH_CSV,
  header: [
    { id: 'model', title: 'MODEL' },
    { id: 'serial', title: 'SERIAL' },
    { id: 'phone', title: 'PHONE' },
    { id: 'customer', title: 'CUSTOMER' },
  ],
  encoding: 'utf8',
});
const records = [];

readXlsxFile(PATH_DATA).then((rows, error) => {
  console.log(rows.length);
  console.log(rows[0]);
  for (let i = 1; i < rows.length - 1; i++) {
    const [requestId, phone, md, sr, mt, date, cus] = rows[i];

    const data = (md + '').replace(/[\r\n\x0B\x0C\u0085\u2028\u2029]+/g, ' ');

    // if (md.indexOf('\n') > -1) {
    //   console.log('data', data);
    // }

    const [prefix, model, serial, sdt] = data.split(' ');
    records.push({ model: model, serial: serial, phone: sdt, customer: cus });
  }
  csvWriter.writeRecords(records).then(() => {
    console.log('Done...');
  });
});
