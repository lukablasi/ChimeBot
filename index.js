const express = require('express');
var cron = require('node-cron');
const playwright = require('playwright');
var request = require('request');
var XLSX = require('xlsx')
var fs = require('fs');
  
const app = express();
const PORT = 3000;

const hook = 'https://hooks.chime.aws/incomingwebhooks/597cf2ec-ddf8-458c-af4c-88bc2adc7a48?token=UzVZd0lZWk18MXxabnlWbDBteHlrMEJjeVdKSDlER2daR283b19fQldmX1pFZVg5eXJkVk5R'

async function pullData() {

  const browser = await playwright.chromium.launch({
      headless: true // setting this to true will not run the UI
  });
  
  const page = await browser.newPage();
  await page.goto('https://fclm-portal.amazon.com/reports/processPathRollup?warehouseId=LCY2');
  await page.click('input[value="Intraday"]');
  await page.locator('select#startHourIntraday').selectOption('7')
  await page.locator('select#startMinuteIntraday').selectOption('30')
  await page.locator('select#endHourIntraday').selectOption('18')
  await page.locator('select#endMinuteIntraday').selectOption('30')
  await page.getByText('HTML').click()
  await page.waitForTimeout(5000);

  const ibTotalVol = (await page.locator('tr', {hasText: 'IB Total'}).locator('.actualVolume').locator('.original').textContent()).slice(1);
  const ibTotalRate = (await page.locator('tr', {hasText: 'IB Total'}).locator('.actualProductivity').locator('.original').textContent()).slice(1);
  const etiRate = (await page.locator('tr', {hasText: 'Each Transfer In - Total'}).locator('.actualProductivity').locator('.original').textContent()).slice(1);
  const stpRate = (await page.locator('tr', {hasText: 'Each Stow to Prime - Total'}).locator('.actualProductivity').locator('.original').textContent()).slice(1);
  const tsiDecantRate = (await page.locator('tr', {hasText: 'Transfer In Decant'}).locator('.actualProductivity').locator('.original').textContent()).slice(1);

  await browser.close();

  return {ibTotalVol, ibTotalRate, etiRate, stpRate, tsiDecantRate}
}

async function downloadExcelFile() {
  let date = new Date();
  let day = ("0" + date.getDate()).slice(-2);
  let month = ("0" + (date.getMonth() + 1)).slice(-2);
  let year = date.getFullYear();

  const browser = await playwright.chromium.launch({
    headless: false // setting this to true will not run the UI
});
  const page = await browser.newPage();
  await page.goto('file://///ant.amazon.com/dept-eu/LCY2/Inbound/Pre%20Shift%20Plan/Days/2023/March');

  const [ download ] = await Promise.all([

    page.waitForEvent('download'),

    page.locator(`text='${day}.${month}.${year} Shift Plan.xlsm'`).click(),
  ]);

  await download.saveAs('D:/projects/flow-bot/plan.xlsm');

  await browser.close();
}

async function readExcelFile() {
  let shiftPlan
  if (fs.existsSync('D:/projects/flow-bot/plan.xlsm')) {
  var file = XLSX.readFile('plan.xlsm');
  let data = []
  const sheets = file.SheetNames
  for (let i = 0; i < sheets.length; i++) {
    const temp = XLSX.utils.sheet_to_json(
      file.Sheets[file.SheetNames[i]])
 temp.forEach((res) => {
    data.push(res)
 })
  }
  shiftPlan = data[6].__EMPTY_3
  shiftPlan = shiftPlan.toLocaleString()
  return shiftPlan
} else{
  shiftPlan = '?';
  return shiftPlan
}
}

async function deleteExcelFile() {
  fs.unlinkSync('D:/projects/flow-bot/plan.xlsm');
}

async function main() {
  const data = await pullData()
  const shiftPlan = await readExcelFile()
  let planDeviation
  let hourlyToPlan;
  if (shiftPlan != '?') {
    let checkTime = new Date();
    let currentHour = checkTime.getHours()
    let currentMinute = checkTime.getMinutes()

    let planModifier;
    
    if (currentHour == 8 && currentMinute == 0) {

    } else if (currentHour == 8 && currentMinute == 30) {
      planModifier = 5
    } else if (currentHour == 9 && currentMinute == 0) {
      planModifier = 11
    } else if (currentHour == 9 && currentMinute == 30) {
      planModifier = 16
    } else if (currentHour == 10 && currentMinute == 0) {
      planModifier = 22
    } else if (currentHour == 10 && currentMinute == 30) {
      planModifier = 27
    } else if (currentHour == 11 && currentMinute == 0) {
      planModifier = 32
    } else if (currentHour == 11 && currentMinute == 30) {
      planModifier = 36
    } else if (currentHour == 12 && currentMinute == 0) {
      planModifier = 40
    } else if (currentHour == 12 && currentMinute == 30) {
      planModifier = 46
    } else if (currentHour == 13 && currentMinute == 0) {
      planModifier = 51
    } else if (currentHour == 13 && currentMinute == 30) {
      planModifier = 56
    } else if (currentHour == 14 && currentMinute == 0) {
      planModifier = 62
    } else if (currentHour == 14 && currentMinute == 30) {
      planModifier = 67
    } else if (currentHour == 15 && currentMinute == 0) {
      planModifier = 72
    } else if (currentHour == 15 && currentMinute == 30) {
      planModifier = 78
    } else if (currentHour == 16 && currentMinute == 0) {
      planModifier = 82
    } else if (currentHour == 16 && currentMinute == 30) {
      planModifier = 86
    } else if (currentHour == 17 && currentMinute == 0) {
      planModifier = 90
    } else if (currentHour == 17 && currentMinute == 30) {
      planModifier = 94
    } else if (currentHour == 18 && currentMinute == 0) {
      planModifier = 98
    } else if (currentHour == 18 && currentMinute == 30) {
      planModifier = 100
    } else {
      planModifier = 100
    }

    hourlyToPlan = Number(data.ibTotalVol.replace(',', '')) - (Number(shiftPlan.replace(',', '') * planModifier / 100))
    hourlyToPlan = Math.round(hourlyToPlan)
    hourlyToPlan = hourlyToPlan.toLocaleString()
    planDeviation = Number(data.ibTotalVol.replace(',', '')) - Number(shiftPlan.replace(',', ''))
    planDeviation = Math.round(planDeviation)
    planDeviation = planDeviation.toLocaleString()
  } else {
    planDeviation = '?'
    hourlyToPlan = '?'
  }
  
  request.post({ headers: {'content-type' : 'application/json'}
               , url: hook
               , body: `{"Content":"/md Metric|Value \\n ---|---| \\n IB Shift Plan:|${shiftPlan}| \\n IB Total:|${data.ibTotalVol}| \\n To total plan:|${planDeviation}| \\n To hourly plan:|${hourlyToPlan}| \\n  ---|---| \\n ETI Rate:|${data.etiRate}| \\n STP Rate:|${data.stpRate}| \\n TSI Decant Rate:|${data.tsiDecantRate}| \\n TPH:|${data.ibTotalRate}|"}` }
               ,function(error, response, body) {
                 console.log(body)
               })
}


cron.schedule('00,30 9-18 * * *', () => main()); 
cron.schedule('50 8,9 * * *', () => downloadExcelFile()); 
cron.schedule('45 18 * * *', () => deleteExcelFile()); 

app.listen(PORT, (error) =>{
    if(!error)
        console.log("Server is successfully running on port "+ PORT)
    else 
        console.log("Error occurred, server can't start", error);
    }
);
