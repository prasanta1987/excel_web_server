const Excel = require('exceljs');
const fs = require('fs')
const express = require('express')


const server = express()
const port = process.env.PORT || 80

var workbook = new Excel.Workbook();

function excelToHtml() {
    workbook.xlsx.readFile('./1.xlsx').then(() => {
        let colCount = workbook.getWorksheet().columnCount
        let rowCount = workbook.getWorksheet().rowCount

        // cellBgColor = workbook.getWorksheet().getCell(1, 1).fill
        // testVal = workbook.getWorksheet().getCell(1, 1).style
        // console.log(testVal)
        let html = `<!DOCTYPE html>
                    <html lang="en">
                    <head>
                        <meta charset="UTF-8">
                        <meta http-equiv="Cache-control" content="no-cache">
                        <meta http-equiv="refresh" content="7">
                        <meta name="viewport" content="width=device-width, initial-scale=1.0">
                        <meta http-equiv="X-UA-Compatible" content="ie=edge">
                        <title>Document</title>
                        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
                      <style>
                            table, tr, td {
                                border: 1px solid black;
                                text-align:center;
                            }
                        </style>
                    </head>
                    <body>
                `

        let body = '<table class="table table-dark">'
        for (row = 1; row <= rowCount; row++) {
            body += '\n'
            body += '<tr>'
            body += '\n'
            for (col = 1; col <= colCount; col++) {
                isFormula = workbook.getWorksheet().getCell(row, col).formulaType
                if (isFormula) {
                    val = workbook.getWorksheet('Sheet1').getCell(row, col).result
                } else {
                    val = workbook.getWorksheet('Sheet1').getCell(row, col).value
                }
                val = (val != null) ? val : '';

                fontColor = workbook.getWorksheet().getCell(row, col).font.color
                fontColor = (!fontColor) ? '00000000' : fontColor.argb;
                bgColor = testVal = workbook.getWorksheet().getCell(row, col).fill
                bgColor = (!bgColor) ? 'FFFFFFFF' : bgColor.fgColor.argb
                body += `<td style="color:${getColor(fontColor)};background-color:${getColor(bgColor)}">${val}</td>`
                // body += `<td>${val}</td>`
                body += '\n'
            }
            body += '</tr>'
        }

        html += body
        html += `</body></html>`
        fs.writeFileSync('./index.html', html)
    })
}

function getColor(color) {
    setColor = '#' + color.substring(2, color.length)
    return setColor
}

setInterval(excelToHtml, 5000)

server.get('*', (req, res) => {
    res.writeHead(200, 'text/html')
    res.end((fs.readFileSync('./index.html')))
})

server.listen(port, () => {
    console.log(`Server Runing on ${port}`)
})