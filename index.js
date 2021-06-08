const axios = require('axios');
const cheerio = require('cheerio');
const moment = require('moment');
const fs = require('fs');
const xlsx = require('xlsx');
const config = require('./config.json');

const CODES = config.LIST_SAHAM;

const MAX_RANK = config.MAX_RANK;

const formatDate = (d) => {
  const [day, month, year] = d.split('/');
  return `${month}/${day}/${year}`;
}

const sleep = (duration) => {
  return new Promise(resolve => {
    setTimeout(() => {
      resolve()
    }, duration * 1000)
  })
}

const FROM_DATE = moment(
  new Date(formatDate(config.FROM_DATE))
);
const TO_DATE = moment(
  new Date(formatDate(config.TO_DATE))
);

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
  
  let timeoutCount = 0;
  
  for (let i = moment(FROM_DATE); i.diff(TO_DATE, 'days') <= 0; i.add(1, 'days')) {
    // it is in the future!
    if (moment().diff(i) < 0) {
      continue;
    }
    
    // saturday or sunday, skip
    if (i.day() === 6 || i.day() === 0) {
      continue;
    }

    let result = []
    
    for (let CODE of CODES) {
      console.log(`Fetching ${i.format('DD-MM-YYYY')} for ${CODE}`);

      // let allBrokerSummary = await getBrokerSummary(CODE, i.format('MM/DD/YYYY'), "all"); // Broker Summary All
      let foreignBrokerSummary = await getBrokerSummary(CODE, i.format('MM/DD/YYYY'), "F"); // Broker Summary Foreign
      let domesticBrokerSummary = await getBrokerSummary(CODE, i.format('MM/DD/YYYY'), "D"); // Broker Summary Domestic

      let charting = await getCharting(CODE, i);
      
      result.push({
        "<ticker>": CODE.toUpperCase(),
        "<date>": i.format('MM/DD/YYYY'),
        "<open>": charting[1],
        "<high>": charting[2],
        "<low>": charting[3],
        "<close>": charting[4],
        // "<volume>": allBrokerSummary.totalBuyLot.value - allBrokerSummary.totalSellLot.value,
        "<volume>": domesticBrokerSummary.totalBuyLot.value - domesticBrokerSummary.totalSellLot.value,
        "<oi>": foreignBrokerSummary.totalBuyLot.value - foreignBrokerSummary.totalSellLot.value,
      })

      timeoutCount++;

      if (timeoutCount >= 75) {
        console.log('Waiting for 5 seconds to prevent IP banned')
        await sleep(5);
        timeoutCount = 0;
      }
    }

    
    fs.writeFileSync(`./out/${i.format('YYYYDDMM')}.txt`, generateCSV(result));
  }

  console.log(">DONE!")
}

const getBrokerSummary = async (code, date, type) => {
  const res = await axios.get(`
    https://www.indopremier.com/module/saham/include/data-brokersummary.php?code=${code.toLowerCase()}&start=${date}&end=${date}&fd=${type}&board=RG
  `);

  const data = await extractDataBrokerSummary(res.data);

  let result = {
    code,
    date,
    ...data
  }

  return result
}

const extractDataBrokerSummary = async (html) => {
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

const getCharting = async (code, date) => {
  const res = await axios.get(`
    https://www.indopremier.com/module/saham/include/json-charting.php?code=${code.toLowerCase()}&start=${moment(date).subtract(1, 'days').format('MM/DD/YYYY')}&end=${date.format('MM/DD/YYYY')}
  `);

  if (res.data) {
    // 25200000 = 7 hours / GMT+7
    return res.data.find(d => d[0] == date.valueOf() + 25200000) || [0, 0, 0, 0, 0, 0]
  }

  return [0, 0, 0, 0, 0, 0]
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

  xlsx.writeFile(workbook, `./${FROM_DATE.format('DD-MM-YYYY')}-${TO_DATE.format('DD-MM-YYYY')}.xlsx`);
}

const generateCSV = (json) => {
  const items = json
  const replacer = (key, value) => value === null ? '' : value // specify how you want to handle null values here
  const header = Object.keys(items[0])
  const csv = [
    header.join(','), // header row first
    ...items.map(row => header.map(fieldName => JSON.stringify(row[fieldName], replacer)).join(','))
  ].join('\r\n')

  return csv.replace(/\"/g, "");
}

main();