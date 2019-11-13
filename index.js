const XLSX = require('xlsx')
const path = require('path')
const https = require('https')

function getStockReport () {
    const dailyReportUrl = 'https://www.twse.com.tw/exchangeReport/STOCK_DAY'
    const dateMonth = '20191101'
    const stockCode = '0056'
    https.get(`${dailyReportUrl}?response=json&date=${dateMonth}&stockNo=${stockCode}`, (res) => {
        res.setEncoding('utf8');
        res.on('data', (chunk) => {
            const resBody = JSON.parse(chunk);
            const fields = resBody.fields
            const data = resBody.data
        });

    }).on("error", (err) => {
        console.log("Error: " + err.message);
    })
}

function init () {
    try {
        let workbook = XLSX.readFile('./0056-orig.xlsx');

        const sheetNames = workbook.SheetNames

        let sheet = workbook.Sheets[sheetNames[0]]

        // console.log('sheet :', sheet);

        // sheet = {...sheet,
        //     'A461': {
        //         t: 's',
        //         v: '108/11/14',
        //         r: '<t>108/11/14</t>',
        //         h: '108/11/14',
        //         w: '108/11/14'
        //     },
        //     'B461': { t: 'n', v: 4689000, w: '4689000 ' },
        //     'C461': { t: 'n', v: 168888888, w: '168888888' },
        //     'D461': { t: 'n', v: 27.3, w: '27.3' },
        //     'E461': { t: 'n', v: 27.2, w: '27.2' },
        //     'F461': { t: 'n', v: 27.1, w: '27.1' },
        //     'G461': { t: 'n', v: 27.8, w: '27.8' },
        //     'H461': { t: 'n', v: 0.06, w: '0.06' },
        //     'I461': { t: 'n', v: 2600, w: '2600' },
        // }

        workbook.Sheets[sheetNames[0]] = sheet;

        XLSX.writeFile(workbook, '0056.xlsx', {type: 'buffer'})
    } catch (err) {}
}


init()