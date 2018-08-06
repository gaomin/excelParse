const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');
const path = require('path');
const parse2html = require('./src/parse2html');

XlsxPopulate.fromFileAsync(path.join(__dirname, "./xlsx/", "in2.xlsx"))
    .then(workbook => {
        var htmlStr = parse2html(workbook.sheet(0));
        outFile(htmlStr)
    });


function outFile(htmlStr) {
    fs.writeFile('./dist/out.html', htmlStr, (err) => {
        if (err) throw err;
        console.log('\n' + 'Your tables have been created');
    });
}
