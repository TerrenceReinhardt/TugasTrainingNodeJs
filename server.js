require('dotenv').config();
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const puppeteer = require('puppeteer');

async function readExcelAndUpload() {
  try {
    const filePath = path.resolve(__dirname, 'TrainingUniAir.xlsx');

    if (!fs.existsSync(filePath)) {
      throw new Error(`Excel file not found at ${filePath}`);
    }

    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });

    worksheet.forEach(row => {
      if (!row.Status) row.Status = '';
    });

    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();

    await page.goto('http://127.0.0.1:8000/users/create', { waitUntil: 'networkidle2' });

    for (let i = 0; i < worksheet.length; i++) {
      const row = worksheet[i];
      const name = row.Name || '';
      const email = row.Email || '';

      if (!name || !email) {
        row.Status = "Failed: Missing required fields";
        continue;
      }

      try {
        console.log(`Uploading user: Name=${name}, Email=${email}`);

        await page.evaluate(() => {
          document.querySelector('input[name="name"]').value = '';
          document.querySelector('input[name="email"]').value = '';
        });

        await page.type('input[name="name"]', name);
        await page.type('input[name="email"]', email);

        await Promise.all([
          page.click('button[type="submit"]'),
          page.waitForSelector('.alert-success', { timeout: 5000 }),
          page.waitForNavigation({ waitUntil: 'networkidle2' }) // Ensure page reloads after form submission
        ]);

        if (page.url().includes('/users')) {
            row.Status = "Success";
          } else {
            row.Status = "Failed: Not redirected to /users";
          }

        let successMessage = null;
        try {
          await page.waitForSelector('.alert-success', { timeout: 10000 });
          successMessage = await page.evaluate(() => {
            const alert = document.querySelector('.alert-success');
            return alert ? alert.textContent.trim() : null;
          });
        } catch (error) {
          console.error(`Timeout waiting for success message on: ${name} (${email})`);
          await page.screenshot({ path: `error_${name}.png` });
        }

        row.Status = successMessage ? "Success" : "Failed: No success message";

        await page.goto('http://127.0.0.1:8000/users/create', { waitUntil: 'networkidle2' });

      } catch (error) {
        console.error(`Error processing ${name}: ${error.message}`);
        await page.screenshot({ path: `error_${name}.png` });
        row.Status = `Failed: ${error.message}`;
      }
    }

    const updatedWorksheet = xlsx.utils.json_to_sheet(worksheet);
    const updatedWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(updatedWorkbook, updatedWorksheet, sheetName);

    const outputFilePath = path.resolve(__dirname, 'TrainingUniAirStatus.xlsx');
    xlsx.writeFile(updatedWorkbook, outputFilePath);
    console.log(`User upload completed. Results saved in ${outputFilePath}`);

    await browser.close();
  } catch (error) {
    console.error('Error processing Excel file:', error.message);
  }
}

readExcelAndUpload();