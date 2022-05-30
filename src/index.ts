import * as cheerio from 'cheerio';
import axios from 'axios';
import ExcelJS from 'exceljs';
import { readFile } from 'fs/promises'
import * as  path from 'path'


const URL = "https://api.cargolargo.com/Cargolargo_MainDetails.aspx?MainID=3535";
const selector = '#form1 .bidSalesResults';

type Lot = {
    lotNumber: number,
    soldFor: number
}

async function main() {
    // const cargoLargoHTMLData = await fetchCargoLargoData();
    const cargoLargoHTMLData = await readFile(path.resolve('CargoLargo_MainDetails.html'), {encoding: 'utf-8'});
    const {lots, resultsDate} = await processData(cargoLargoHTMLData);
    generateSpreadsheet(lots, 
        `${resultsDate}_bid_results`,

        `${resultsDate}_bid_results.xlsx`,
        )


}



async function processData(data: any) {
    const $ = cheerio.load(data)

    const resultsDate = $('h3').text();
    const bidSalesResultsTable = $('.bidSalesResults')
    const rows = $('tr', bidSalesResultsTable)
    const lots: Lot[] = []

    rows.each((i, el) => {
        const lot: Lot = {
            lotNumber: Number($(el).find('.lotNumber').text()),
            soldFor: Number($(el).find('.soldFor').text().replace(/[^0-9.-]+/g, ""))
        }

        lot.lotNumber && lots.push(lot)
    })

    const formattedDate = resultsDate.slice(resultsDate.indexOf(': ') + 2)
        .split('').filter(chr => chr !== '/').join('').padStart(8, '0')

    return { resultsDate: formattedDate, lots }
}

async function generateSpreadsheet(data: Lot[], worksheetName: string, filename: string) {

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Brandon Zamorano';
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet(worksheetName)

    //Define a list a columns
    const headerColumnStyle: Partial<ExcelJS.Style> = {
        fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: {
                argb: "333E4F"
            }
        },
        font: {
            bold: true,
            color: {
                argb: 'FFFFFF'
            }
        }
    }
    const numFmtStr = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';

    const columnHeaders: Partial<ExcelJS.Column>[] = [
        {
            header: "Lot Number",
            key: 'ln',
            width: 11,
            alignment: { horizontal: 'left' },
            style: headerColumnStyle
        },
        {
            header: "Sold For",
            key: 'sf',
            width: 14,
            alignment: { horizontal: 'left' },
            style: { ...headerColumnStyle, numFmt: numFmtStr }

        }
    ]

    worksheet.columns = columnHeaders;
    const lotDataRows = worksheet.addRows(data.map(lot => [lot.lotNumber, lot.soldFor]))
    lotDataRows.forEach((row, idx) => {
        row.eachCell((cell, cellNum) => {
            cell.style.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: {
                    argb: "E3E3E3"
                }
            }
            cell.font = { color: { argb: "000000" } }
            cell.border = {
                top: {
                    style: 'thin',
                    color: {
                            argb: "333333"
                        }
                    }
                }
            
            if(cellNum === 1) {
            cell.border = {
                ...cell.border,
                right: {
                    style: 'thin'
                }
            }
            cell.alignment = { horizontal: 'left' }
        }
    })
})

worksheet.views = [
    { state: 'frozen', ySplit: 1 }
]
worksheet.autoFilter = `A1:B${worksheet.actualRowCount}`

await workbook.xlsx.writeFile(filename)
    
}

main()