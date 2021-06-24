const readXlsxFile = require('read-excel-file/node');
const axios = require('axios');
const moment = require('moment');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;

const PATH_DATA = './data/data_20210614.xlsx';
const PATH_CSV = './out/data_20210614_activated.csv';

const agencies = {
  'Điện Máy Lạc Hồng': 159,
  Rindo: 159,
  'Điện Máy Quyết Chi': 269,
  'Trần Văn Toan': 267,
  'Điện Máy Minh Thu': 268,
  'Điện Máy Hải Linh': 136,
  'Điện Máy Hải Ninh': 136,
  'Điện Máy Nam Hạnh': 266,
  'Điện máy Phú Lý': 260,
  'Điện máy Phúc Lý': 260,
  'Cửa Hàng Sinh NGọc': 168,
  'Điện Máy  Khắc Lượng': 237,
  'Điện Lạnh Hải Lâm': 165,
  'Cửa hàng điện tử Tiến Dũng': 206,
  'ĐL Đức Thành': 189,
  'CH Hoàn Anh': 121,
  'ĐL Trung Đức': 82,
  'Đại lý Vinh Kiều': 166,
  'điện tử Thao Cúc': 163,
  'Điện Máy Ngọc Toản': 124,
  'Điện Máy Tuấn  Dung': 188,
  'Điện Máy An Viên': 235,
  'Điện Máy Tân Thủy': 151,
  'Điện Máy Thành Chung': 239,
  'CH Cường Mai': 125,
  'Điện Máy Hồng Hảo': 32,
  'Điện Máy Huân Hiền': 175,
  'Điện Máy Chung Anh': 253,
  'Điện Máy Chiến Hiền': 102,
  'CH Huy Xuyến': 157,
  'Điện Máy Hải Âu': 176,
  'Điện Máy Thảo Phú': 225,
  'Điện Máy Phú Thiết': 178,
  'Điện Máy Thân Hương': 201,
  'Điện Máy Quang Trung': 105,
  'Điện Máy Hà Quyên': 62,
  'Điện Máy Hoàng Quân': 179,
  'Điện Máy  Bình Lâm': 152,
  'Điện máy Cường Mai': 31,
  'Điện Máy Chiến Chảo': 241,
  'Điện Máy Quang Hương': 238,
  'Điện máy Thống Ngân': 210,
  'Điện Máy Hải Thủy': 231,
  'Nguyễn Đức Bảo': 77,
  'Điện Máy Thành Hiền': 229,
  'Điện Máy Tất Thắng': 214,
  'CH Vinh Hoa': 33,
  'Điện tử Hậu Hà': 43,
  'Điện Máy Uỷ Nam': 213,
  'Điện Máy Thắng Thanh': 185,
  'Điện Máy Quang Minh 2': 196,
  'Điện Máy Quyến Trang': 61,
  'Điện máy Ngọc Giang': 227,
  'Điện Máy Khương Loan': 234,
  'Điện Máy Việt Phương': 204,
  'Điện Máy Văn Mười': 174,
  'Cửa hàng thế giới bếp HS': 16,
  'Điện Máy Diên Hương': 205,
  'Điện máy Việt Thám': 129,
  'Điện Máy Hoàng Anh': 215,
  'Điện máy Đồng Tâm': 54,
  'ĐL Đông Ngâm': 41,
  'Điện Máy Tiến Cúc': 67,
  'Điện Máy  Quân Kon Tum': 216,
  'ĐL Huy Khôi': 9,
  'Điện Máy Nguyễn Văn Dương': 251,
  'Điện máy Hương Quân': 52,
  'Điện máy Nghị Mai': 139,
  'Tuyên Ngần': 71,
  'Điện Máy Ngọc Huyền': 173,
  'Điện máy Vinh Quỳnh': 128,
  'Điện Máy Hiền Khản': 170,
  'Điện Máy Đức Chính': 162,
  'Điện máy Đồi Dung': 137,
  'Điện Máy Khánh Huyền': 198,
  'Điện Máy Khánh Linh': 126,
  'Điện Máy Hải Vân': 232,
  'Điện máy Hùng Hiên': 220,
  'Điện Máy Hưng Hải': 112,
  'Điện Máy Huy Ba': 187,
  'Điện Máy Thành Đông': 68,
  'Điện Máy Cầm Phát': 127,
  'Điện Máy Việt Cường': 218,
};

const searchAgency = async (name) => {
  const res = await axios.get(
    encodeURI(
      `http://baohanh.rindo.vn:8081/api/agencies?q=${name}&sort=id,ASC&page=0&size=1`
    ),
    {
      headers: {
        Authorization:
          'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6MSwidXNlcm5hbWUiOiJhZG1pbiIsImF1dGhvcml0aWVzIjpbIlJPTEVfQURNSU4iLCJST0xFX1NVUEVSX1VTRVIiLCJST0xFX1VTRVIiXSwiaWF0IjoxNjE3ODU1NzU3LCJleHAiOjE2MTc4NTYwNTd9.kT3tYMtb7feHhj5Ajk2bwSh5URXwqKxZbje3vT6EmE4',
      },
    }
  );
  const agency = res.data[0];
  if (!agency) {
    throw new Error(`SearchAgency not found: ${name}`);
  }
  agencies[name] = agency.id;
  return agency;
};

const workActivatedDevices = async () => {
  const csvWriter = createCsvWriter({
    path: PATH_CSV,
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
  const agenciesNotFound = [];

  readXlsxFile(PATH_DATA).then(async (rows, error) => {
    console.log(`Total ${rows.length} rows.\nHeader are: ${rows[0]}`);
    for (let i = 1; i < rows.length - 1; i++) {
      const [requestId, phone, md, sr, mt, date, _cus] = rows[i];
      const cus = (_cus || '').trim();

      // replace all newLines characters to space
      const data = (md + '').replace(/[\r\n\x0B\x0C\u0085\u2028\u2029]+/g, ' ');

      const [prefix, model, serial, sdt] = data.split(' ');

      if (!agencies[cus] && !agenciesNotFound[cus]) {
        agenciesNotFound.push(cus);
      }
      records.push({
        index: i,
        activationCode: serial,
        model: model,
        produceDate: moment(date).format('YYYYMMDD'),
        agencyId: agencies[cus] || '159',
        monthsOfActivation: 24,
      });
    }
    csvWriter.writeRecords(records).then(() => {
      console.log(`======\nSuccessful!\nSaved to: ${PATH_CSV}\n==========`);
    });

    console.log(
      '====\nAgency not found in Excel file\n',
      agenciesNotFound.join('\n')
    );
  });
};

const main = async () => {
  // find missing agency id
  const keys = Object.keys(agencies);
  for (let i = 0; i < keys.length; i++) {
    try {
      const name = keys[i];
      if (!agencies[name] || agencies[name] == 9999) {
        const agency = await searchAgency(name);
        agencies[name] = agency.id;
      }
    } catch (err) {
      console.log(err.message);
    }
  }
  // console.log('agencies', JSON.stringify(agencies, null, 2));
  await workActivatedDevices();
};

module.exports = main;
