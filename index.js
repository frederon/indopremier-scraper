const axios = require('axios');
const cheerio = require('cheerio');
const moment = require('moment');
const fs = require('fs');
const xlsx = require('xlsx');
const config = require('./config.json');

const CODES = config.LIST_SAHAM;
const MONTH_YEAR = config.MONTH_YEAR;

const MAX_RANK = 5;

const si = [
  { value: 1E3, symbol: "k" },
  { value: 1E6, symbol: "M" },
  { value: 1E9, symbol: "B" },
  { value: 1E12, symbol: "T" },
  { value: 1E15, symbol: "P" },
  { value: 1E18, symbol: "E" }
];

const main = async () => {
  if (!CODES.length) {
    console.log('Please enter CODE name');
    return;
  }

  let result = []

  const days_count = moment(
    new Date(`${MONTH_YEAR[0]}/01/${MONTH_YEAR[1]}`)
  ).daysInMonth();

  for (let CODE of CODES) {
    for (let i = 1; i <= days_count; i++) {
      const date = moment(
        new Date(`${MONTH_YEAR[0]}/${i}/${MONTH_YEAR[1]}`)
      )
  
      // it is in the future!
      if (moment().diff(date) < 0) {
        continue;
      }
      
      // saturday or sunday, skip
      if (date.day() === 6 || date.day() === 0) {
        continue;
      }
  
      let dateResult = await startProcess(CODE, date.format('MM/DD/YYYY'));
  
      result.push(dateResult)
    }
  }

  fs.writeFileSync(`./${MONTH_YEAR[0]}-${MONTH_YEAR[1]}.json`, JSON.stringify(result));

  generateExcel(result);

  console.log(">DONE!")  

  return result;
}

const startProcess = async (code, date) => {
  const res = await axios.get(`
    https://www.indopremier.com/module/saham/include/data-brokersummary.php?code=${code.toLowerCase()}&start=${date}&end=${date}&fd=all&board=all
  `);

  const data = await extractData(res.data);

  let result = {
    code,
    date,
    ...data
  }

  return result
}

const extractData = async (html) => {
  const $ = cheerio.load(html);

  let data = {
    buyer: [],
    seller: [],
    totalBuyLot: {},
    totalSellLot: {},
  }

  $('.table tbody tr').each((rowIndex, row) => {
    let buyRes = {}
    let sellRes = {}

    if (rowIndex <= MAX_RANK - 1) {
      $(row).children().each((colIndex, col) => {
        switch (colIndex) {
          case 0:
            buyRes['name'] = $(col).text()
          case 1:
            buyRes['lot'] = $(col).text()
          case 2:
            buyRes['val'] = $(col).text()
          case 3:
            buyRes['avg'] = $(col).text()
          case 5:
            sellRes['name'] = $(col).text()
          case 6:
            sellRes['lot'] = $(col).text()
          case 7:
            sellRes['val'] = $(col).text()
          case 8:
            sellRes['avg'] = $(col).text()
        }
      });
  
      data.buyer.push(buyRes)
      data.seller.push(sellRes)
    }

  })

  data.totalBuyLot.value = calculateTotalLot(data.buyer)
  data.totalBuyLot.formatted = formatNumber(data.totalBuyLot.value, 2)
  data.totalSellLot.value = calculateTotalLot(data.seller)
  data.totalSellLot.formatted = formatNumber(data.totalSellLot.value, 2)

  return data;
}

const calculateTotalLot = (data) => {
  return data.reduce((sum, d) => {
    let numString = d.lot.replace(/\s/g, '').replace(',', '')

    let unit = /[A-Z]|[a-z]/g.exec(numString);
    let numRes = Number(numString.replace(/[A-Z]|[a-z]/g, ''))

    if (unit && unit[0]) {
      let multiplier = si.find(val => val.symbol === unit[0]).value;
      numRes = numRes * multiplier;
    }

    return sum + numRes
  }, 0);
}

const formatNumber = (num, digits) => {
  let rx = /\.0+$|(\.[0-9]*[1-9])0+$/;
  let i;
  for (i = si.length - 1; i > 0; i--) {
    if (num >= si[i].value) {
      break;
    }
  }
  return (num / si[i].value).toFixed(digits).replace(rx, "$1") + si[i].symbol;
}

const generateExcel = (data) => {
  data = data.map(val => ({
    'Saham': val.code.toUpperCase(),
    'Tanggal': moment(new Date(val.date)).format('DD/MMM/YYYY'),
    [`Net ${MAX_RANK} Buy`]: val.totalBuyLot.value,
    [`Net ${MAX_RANK} Buy Formatted`]: val.totalBuyLot.formatted,
    [`Net ${MAX_RANK} Sell`]: val.totalSellLot.value,
    [`Net ${MAX_RANK} Sell Formatted`]: val.totalSellLot.formatted,
  }));

  const workbook = xlsx.utils.book_new();

  for (let CODE of CODES) {
    const filteredCodeData = data.filter(val => val['Saham'] === CODE.toUpperCase())
    const sheets = xlsx.utils.json_to_sheet(filteredCodeData);
    xlsx.utils.book_append_sheet(workbook, sheets, CODE);
  }

  xlsx.writeFile(workbook, `./${MONTH_YEAR[0]}-${MONTH_YEAR[1]}.xlsx`);
}

main();