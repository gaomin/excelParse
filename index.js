const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');
const path = require('path');
const parse2html = require('./src/parse2html');
const conf = require('./src/config');

XlsxPopulate.fromFileAsync(path.join(__dirname, "./xlsx/", conf.inFile))
    .then(workbook => {
        var htmlStr = parse2html(workbook.sheet(0));
        outFile(htmlStr)
    });


function outFile(htmlStr) {
    fs.writeFile('./dist/'+conf.outFile, htmlStr, (err) => {
        if (err) throw err;
        console.log('\n' + 'Your tables have been created');
    });
}
