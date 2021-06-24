const readXlsxFile = require('read-excel-file/node');
const Excel = require('exceljs');
const axios = require('axios');
const moment = require('moment');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;

const PATH_INPUT = './device_inactive.xlsx'; // XLSX (2004) format
const PATH_OUTPUT = './device_inactive.csv';

const main = async () => {
  const csvWriter = createCsvWriter({
    path: PATH_OUTPUT,
    header: [
      { id: 'index', title: 'index' },
      { id: 'activationCode', title: 'activationCode' },
      { id: 'qrcode', title: 'qrcode' },
      { id: 'model', title: 'model' },
      { id: 'produceDate', title: 'produceDate' },
      { id: 'macAddress', title: 'macAddress' },
      { id: 'agencyId', title: 'agencyId' },
      { id: 'monthsOfActivation', title: 'monthsOfActivation' },
      { id: 'note', title: 'note' },
    ],
    encoding: 'utf8',
  });
  const records = [];

  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(PATH_INPUT);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('Worksheet not found');
  }
  for (let i = 1; i < worksheet.rowCount; i++) {
    const values = worksheet.getRow(i).values;
    const [ignore, serial, model] = values;

    records.push({
      index: i,
      activationCode: serial,
      model: model,
      produceDate: moment().format('YYYYMMDD'),
      monthsOfActivation: 24,
    });
  }

  csvWriter.writeRecords(records).then(() => {
    console.log(`======\nSuccessful!\nSaved to: ${PATH_OUTPUT}\n==========`);
  });
};

main();
