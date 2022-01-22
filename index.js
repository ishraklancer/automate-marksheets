console.log('hello world')

const puppeteer = require('puppeteer');
const url = 'https://mcam.mist.ac.bd/Security/Login.aspx'
// const url = 'https://en.wikipedia.org/wiki/Main_Page'

const url1 = 'https://mcam.mist.ac.bd/Module/result/ExamResultSubmit.aspx?mmi=41505a2c74524d494e63'
const input = {
  username: '204-0030-1',
  password: '##################'
};
const excel = require('xlsx')


var marksIdHead = '#ctl00_MainContainer_GvExamMarkSubmit_ctl'
var marksIdTail = '_txtMark'
var marksIdSerial = 02;
var xlId = 0;
var absentIdHead = '#ctl00_MainContainer_GvExamMarkSubmit_ctl'
var absetIdTail = '_chkStatus'

const login = async () => {
  const browser = await puppeteer.launch({
    headless: false, ignoreHTTPSErrors: true,
    args: [`--window-size=1800,900`],
    defaultViewport: {
      width: 1800,
      height: 900
    }
  });
  const page = await browser.newPage();
  await page.setDefaultNavigationTimeout(0);
  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await page.type('#logMain_UserName', input.username)
  await page.type('#logMain_Password', input.password)
  await page.click('#logMain_Button1');
  await page.waitForNavigation();
  await page.goto(url1, { waitUntil: 'domcontentloaded' });
  await page.select('#ctl00_MainContainer_ucProgram_ddlProgram', '2')
  await page.waitForTimeout(10000)
  await page.select('#ctl00_MainContainer_ddlCourse', '499_1_8599')
  await page.waitForSelector('#divProgress', { hidden: true, timeout: 100000 });
  await page.select('#ctl00_MainContainer_ddlExam', '700')
  console.log('waiting')
  await page.waitForSelector('#divProgress', { hidden: true, timeout: 100000 });
  await page.waitForTimeout(10000)

  console.log('data loaded')

  const formData = await page.evaluate(() => {
    return Array.from(
      document.querySelectorAll('table[id="ctl00_MainContainer_GvExamMarkSubmit"] > tbody > tr'),
      row => Array.from(row.querySelectorAll('td'), cell => cell.innerText)
    )
  })

  // const formData = await page.$$eval('#ctl00_MainContainer_GvExamMarkSubmit tr', rows => {
  //   return Array.from(rows, row => {
  //     const columns = row.querySelectorAll('td');
  //     return Array.from(columns, column => column.innerText);
  //   });
  // });

  // ctl00_MainContainer_GvExamMarkSubmit_ctl02_txtMark

  // console.log(formData)

  const xlData = readExcel();

  for (let i = 0; i < formData.length - 2; i++) {
    var serial = (marksIdSerial < 10) ? "0" + JSON.stringify(marksIdSerial) : JSON.stringify(marksIdSerial)
    var marksId = marksIdHead + serial + marksIdTail;
    var absentId = absentIdHead + serial + absetIdTail;
    
    await page.type(marksId, JSON.stringify(xlData[xlId]['CT - 2\n(20)']))

    if (xlData[xlId]['CT - 2\n(20)'] === 'A') {
      await page.click(absentId)
      // await page.$eval('input[id="name_check"]', check => check.checked = true);

    }
    xlId++;
    marksIdSerial++;
  }

  await page.click('#ctl00_MainContainer_GvExamMarkSubmit_ctl01_SubmitAllMarkButton')

  await page.screenshot({ path: 'screenshot.png', fullPage: true });
  await browser.waitForTarget(() => false)

  await page.click('#ctl00_lblLogout')
  await page.waitForNavigation();

  await browser.close();
}

// puppt()


const readExcel = () => {
  var workbook = excel.readFile('CSE-313-2019-Marks.xlsx')
  var sheetNames = workbook.SheetNames;
  // console.log(sheetNames)
  var xlData = excel.utils.sheet_to_json(workbook.Sheets[sheetNames[3]]);
  // console.log(xlData)
  return xlData
}

login();
