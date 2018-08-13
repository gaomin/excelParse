module.exports = function parse2html(workbook) {
    const columnNum = workbook.row(1)._cells.length;
    const rowNum = workbook._rows.length;
    const mergeCells = workbook._mergeCells;
    const { rowMergeCells, columnMergeCells } = _countMergeObj(workbook, mergeCells);
    const htmlStr = _createHtml(workbook, columnNum, rowNum, rowMergeCells, columnMergeCells);

    return htmlStr;
}

function _countMergeObj(workbook, mergeCells) {
    let rowMergeCells = {};
    let columnMergeCells = {};

    Object.keys(mergeCells).forEach((item, index) => {
        let rangeObj = workbook.range(item);

        if (rangeObj._numRows == 1) {               //合并列
            if (!rowMergeCells[rangeObj._minRowNumber]) {
                rowMergeCells[rangeObj._minRowNumber] = {};
            }
            rowMergeCells[rangeObj._minRowNumber][rangeObj._minColumnNumber] = rangeObj._numColumns;
        } else if (rangeObj._numColumns == 1) {      //合并行
            if (!columnMergeCells[rangeObj._minColumnNumber]) {
                columnMergeCells[rangeObj._minColumnNumber] = {};
            }
            columnMergeCells[rangeObj._minColumnNumber][rangeObj._minRowNumber] = rangeObj._numRows;
        } else {                                    //同时合并行和列, 特殊处理columnMergeCells, rowMergeCells不做特殊处理
            if (!rowMergeCells[rangeObj._minRowNumber]) {
                rowMergeCells[rangeObj._minRowNumber] = {};
            }
            rowMergeCells[rangeObj._minRowNumber][rangeObj._minColumnNumber] = rangeObj._numColumns;

            for (let i = rangeObj._minColumnNumber; i <= rangeObj._maxColumnNumber; i++) {
                if (!columnMergeCells[i]) {
                    columnMergeCells[i] = {};
                }
                columnMergeCells[i][rangeObj._minRowNumber] = rangeObj._numRows;
            }
        }
    })
    return { rowMergeCells, columnMergeCells }
}

function _createHtml(workbook, columnNum, rowNum, rowMergeCells, columnMergeCells) {
    var htmlStr = '\
        <html>\
            <head>\
                <meta http-equiv="Content-Type" Content="text/html; Charset=UTF-8">\
                <meta name="viewport" content="width=device-width, user-scalable=yes, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0">\
            </head>\
            <body>\
                <table style="border-collapse: collapse; white-space: nowrap;" width="100%" border="0" cellpadding="0" cellspacing="0" >\n<tbody>\n';
    let deleteCell = [];
    let deleteColum = [];

    for (let row = 1; row < rowNum; row++) {
        htmlStr += '<tr>\n';

        for (let column = 1; column < columnNum; column++) {
            let styleStr = _createStyle(workbook, row, column);

            if ((rowMergeCells[row] && rowMergeCells[row][column]) && (columnMergeCells[column] && columnMergeCells[column][row])) {
                htmlStr += '<td colspan="' + rowMergeCells[row][column] + '" rowspan="' + columnMergeCells[column][row] + '" style="' + styleStr + '"><span>' + (workbook.row(row).cell(column).value() || '') + '</span></td>'
                preColumn = column;
                column += rowMergeCells[row][column] - 1;

                for (let i = preColumn; i <= column; i++) {
                    for (let j = row + 1; j < row + columnMergeCells[i][row]; j++) {
                        deleteCell.push(JSON.stringify([j, i]));
                    }
                }
            } else if (rowMergeCells[row] && rowMergeCells[row][column]) {
                htmlStr += '<td colspan="' + rowMergeCells[row][column] + '" style="' + styleStr + '">\
                                ' + (workbook.row(row).cell(column).value() || '') + '\
                            </td>';
                column += rowMergeCells[row][column] - 1;
            } else if (columnMergeCells[column] && columnMergeCells[column][row]) {
                htmlStr += '<td rowspan="' + columnMergeCells[column][row] + '" style="' + styleStr + '">\
                                ' + (workbook.row(row).cell(column).value() || '') + '\
                            </td>';
                for (let i = row + 1; i < row + columnMergeCells[column][row]; i++) {
                    deleteCell.push(JSON.stringify([i, column]));
                }
            } else if (deleteCell.indexOf(JSON.stringify([row, column])) > -1) {

            } else {
                htmlStr += '<td style="' + styleStr + '">\
                                ' + (workbook.row(row).cell(column).value() || '') + '\
                            </td>';
            }

        }
        htmlStr += '</tr>\n';
    }

    htmlStr += '</tbody>\n</table>\n</body></html>';

    return htmlStr;
}

function _createStyle(workbook, row, column) {

    let cell = workbook.row(row).cell(column);
    let styleStr = "";

    let bold = cell.style('bold');
    let italic = cell.style('italic');
    let underline = cell.style('underline');
    let strikethrough = cell.style('strikethrough');
    let fontSize = cell.style('fontSize');
    let fontFamily = cell.style('fontFamily');
    let fontColor = cell.style('fontColor');
    let horizontalAlignment = cell.style('horizontalAlignment');
    let fill = cell.style('fill');
    let border = cell.style('border');
    let borderStyle = cell.style('borderStyle');

    if (bold) {
        styleStr += "font-weight: bold;";
    }

    if (italic) {
        styleStr += "font-style: italic;";
    }

    if (underline) {
        styleStr += "text-decoration: underline;"
    }

    if (strikethrough) {
        styleStr += 'text-decoration: line-through';
    }

    if (horizontalAlignment) {
        styleStr += "text-align: " + horizontalAlignment + ";";
    }

    if (fill && fill.color && fill.color.rgb && fill.color.rgb !== 'FFFFFF') {
        if (fill.color.rgb.length > 6) {
            styleStr += "background-color: #" + fill.color.rgb.slice(2) + ";";
        } else {
            styleStr += "background-color: #" + fill.color.rgb + ";";
        }
    }

    if (fill && fill.color && fill.color.theme) {
        styleStr += "background-color: #f00;";
    }

    if (fontColor && fontColor.rgb && fontColor.rgb !== '000000') {
        if (fontColor.rgb.length > 6) {
            styleStr += "color: #" + fontColor.rgb.slice(2) + ";";
        } else {
            styleStr += "color: #" + fontColor.rgb + ";";
        }
    }

    if (fontSize) {
        styleStr += "font-size: " + fontSize + ";";
    }

    if (fontFamily) {
        styleStr += "font-family: " + fontFamily + ";";

    }

    if (border) {
        
        Object.keys(border).forEach((item, index) => {
            if (border[item]) {
                if (border[item].color && border[item].color.rgb) {
                    styleStr += "border-" + item + ': ' + _getborderStyle(borderStyle[item]) + ' #' + border[item].color.rgb + ";";
                }else {
                    styleStr += "border-" + item + ': ' + _getborderStyle(borderStyle[item]) + ' #000' + ";";
                }
            }
        })

    }

    styleStr += "height: 24px;"
    return styleStr;
}

function _getborderStyle (str){
    switch (str) {
        case 'thin':
            return '1px solid'
            break;
        case 'medium':
            return '1px solid'
            break;
        case 'dashed':
            return '1px dashed'
            break;

        default:
            return '1px solid'
            break;
    }
}

