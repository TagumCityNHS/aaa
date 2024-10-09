const inquirer = require('@inquirer/prompts');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { NFC } = require('nfc-pcsc');

const readXLSX = (filePath) => {
  return new Promise((resolve, reject) => {
    if (!fs.existsSync(filePath)) {
      reject(new Error('XLSX file does not exist.'));
      return;
    }

    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    console.table(data);
    resolve({ workbook, data });
  });
};

const askForXLSXFile = async () => {
  const fileName = await inquirer.input({
    message: 'What is the XLSX File Name? (including .xlsx)',
    validate: (input) => {
      const filePath = path.join(process.cwd(), input);
      if (fs.existsSync(filePath) && path.extname(filePath) === '.xlsx') {
        return true;
      } else {
        return 'File does not exist or is not an XLSX file. Please enter a valid file name.';
      }
    },
  });

  return path.join(process.cwd(), fileName);
};

const askForConfirmation = async () => {
  const confirmation = await inquirer.select({
    message: 'Press Y to Continue or X to Exit',
    choices: [
      { name: 'Y', value: 'yes' },
      { name: 'X', value: 'no' },
    ],
  });

  return confirmation;
};

const saveToExcel = (workbook, data, fileName) => {
  const newWorksheet = xlsx.utils.json_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;
  xlsx.writeFile(workbook, fileName);
  console.log('Excel file updated after each scan.');
};

const main = async () => {
  try {
    const fileName = await askForXLSXFile();
    console.log('File Path:', fileName);

    const { workbook, data } = await readXLSX(fileName);
    const nfc = new NFC();

    nfc.on('reader', (reader) => {
      console.log(`${reader.reader.name} device attached`);

      (async () => {
        while (true) {
          const confirmation = await askForConfirmation();

          if (confirmation === 'yes') {
            for (const row of data) {
              const { firstname: firstName, lastname: lastName } = row;
              console.log(`Please Scan ${firstName} ${lastName}`);

              // Wait for the card scan
              await new Promise((resolve) => {
                reader.once('card', (card) => {
                  console.log(`Scanned UID: ${card.uid}`);
                  // Update the current row with the scanned UID (SecuredLRN)
                  row.SecuredLRN = card.uid;
                  // Save the updated data to Excel immediately after each scan
                  saveToExcel(workbook, data, fileName);
                  resolve(); // Resolve the promise after a scan
                });
              });
            }

            // Break the loop after scanning all rows
            console.log('All scans completed.');
            break;
          } else {
            console.log('Operation cancelled.');
            break; // Exit if the user chooses not to continue
          }
        }

        reader.close();
      })();

      reader.on('error', (err) => {
        console.log(`${reader.reader.name} an error occurred`, err);
      });

      reader.on('end', () => {
        console.log(`${reader.reader.name} device removed`);
      });
    });

    nfc.on('error', (err) => {
      console.log('An error occurred', err);
    });
  } catch (error) {
    console.error('Error:', error.message);
  }
};

main();
