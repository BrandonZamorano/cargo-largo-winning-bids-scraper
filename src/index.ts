import puppeteer from 'puppeteer';
import {read, writeFileXLSX, utils, writeFile} from 'xlsx'

const URL = "https://api.cargolargo.com/Cargolargo_MainDetails.aspx?MainID=3535";
const selector = '#form1 .bidSalesResults';



(async () => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(URL);
    await page.waitForSelector(selector, {timeout:10000});
    // await page.screenshot({path:'./screenshot2.png'})
    const content = await page.content();
    const lots = await page.evaluate(() => {
        const [, ...rows] = [...document.querySelectorAll('.bidSalesResults tr')]
        return rows.map(row => {
            const lot = {
                lotNumber: row.children[0].textContent,
                soldFor: row.children[1].textContent
            }
            return lot
        })
    })

    await browser.close()

    const worksheet = utils.json_to_sheet(lots);
    utils.sheet_add_aoa(worksheet, [["Lot Number", "Price Sold For"]], {origin: "A1"})
    const workbook = utils.book_new();
    utils.book_append_sheet(workbook, worksheet, `052722 Winning Bid Sales`);

    // console.log("Content: ", content)
    // console.log(`Lots: ${JSON.stringify(lots)}`)
    lots.forEach(lot => console.log(lot))
    writeFile(workbook, "052722_Cargo_Largo_Winning_Bid_Sales.xlsx")
    
})()