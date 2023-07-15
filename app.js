const express = require("express");
const app = express();
const { readFile } = require("xlsx");
const cors = require("cors");

app.use(cors({}));

let rukuhs = [];

app.get("/data", (req, res) => {
  const input = req.query.input; // Assuming the input is passed as a query parameter

  // Load the Excel file
  const workbook = readFile("all-Distribuation.xlsx");

  // Get the "Database" sheet
  const databaseSheet = workbook.Sheets["Database"];

  // Extract the surah number and ayah range from the input
  const regex = /\((\d+):(\d+)\)\((\d+):(\d+)\)/;
  const [, surahStart, ayahStart, surahEnd, ayahEnd] = input.match(regex);

  // Convert the values to numbers
  const surahStartNum = parseInt(surahStart, 10);
  const ayahStartNum = parseInt(ayahStart, 10);
  const surahEndNum = parseInt(surahEnd, 10);
  const ayahEndNum = parseInt(ayahEnd, 10);

  // Iterate over the rows in the "Database" sheet and filter the relevant data
  const filteredData = [];
  const range = databaseSheet["!ref"].split(":");
  const startRow = parseInt(range[0].replace(/\D/g, ""), 10);
  const endRow = parseInt(range[1].replace(/\D/g, ""), 10);

  let isSurahStarted = false; // Flag to indicate if the surah has started
  for (let i = startRow; i <= endRow; i++) {
    const surahName = databaseSheet[`F${i}`].v;
    const rowSurahNo = databaseSheet[`E${i}`].v;
    const rowAyahNo = databaseSheet[`K${i}`].v;
    const rowAyahText = databaseSheet[`M${i}`].v;

    if (
      !isSurahStarted &&
      rowSurahNo === surahStartNum &&
      rowAyahNo >= ayahStartNum
    ) {
      isSurahStarted = true; // Mark the start of the surah
    }

    if (isSurahStarted) {
      filteredData.push({
        surahName: surahName,
        surahNo: rowSurahNo,
        ayahNo: rowAyahNo,
        ayahText: rowAyahText,
      });

      if (rowSurahNo === surahEndNum && rowAyahNo === ayahEndNum) {
        break; // Stop iterating after reaching the end ayah of the last specified surah
      }
    }
  }

  res.json(filteredData);
});

app.get("/juz", (req, res) => {
  const juzValue = req.query.juz; // Assuming the juz value is passed as a query parameter

  // Load the Excel file
  const workbook = readFile("all-Distribuation.xlsx");

  // Get the "Database" sheet
  const databaseSheet = workbook.Sheets["Database"];

  // Iterate over the rows in the "Database" sheet and filter the relevant data by juz value
  const filteredData = [];
  const range = databaseSheet["!ref"].split(":");
  const startRow = parseInt(range[0].replace(/\D/g, ""), 10);
  const endRow = parseInt(range[1].replace(/\D/g, ""), 10);

  for (let i = startRow; i <= endRow; i++) {
    const rowJuzValue = databaseSheet[`B${i}`].v;

    if (rowJuzValue == juzValue) {
      const surahNo = databaseSheet[`E${i}`].v;
      const surahName = databaseSheet[`F${i}`].v;
      const ayahNo = databaseSheet[`K${i}`].v;
      const ayahText = databaseSheet[`M${i}`].v;
      filteredData.push({ surahNo, surahName, ayahNo, ayahText });
    }
  }

  res.json(filteredData);
});

function getValueFromSheet(req, res, sheetName, sheetNo) {
  const parahNo = req.query.parahNo;
  const valueNo = req.query[sheetNo];

  // Load the Excel file
  const workbook = readFile("all-Distribuation.xlsx");

  // Get the specified sheet
  const sheet = workbook.Sheets[sheetName];

  const columnLetter = String.fromCharCode(65 + parseInt(valueNo)); // Convert valueNo to column letter (A = 1, B = 2, etc.)
  const cellAddress = `${columnLetter}${parseInt(parahNo) + 1}`;
  const value = sheet[cellAddress]?.v;

  if (value) {
    res.json({ ayahNo: value });
  } else {
    res.status(404).json({ error: `${sheetName} value not found` });
  }
}

app.get("/ruba", (req, res) => {
  getValueFromSheet(req, res, "Ruba", "rubaNo");
});

app.get("/nisf", (req, res) => {
  getValueFromSheet(req, res, "nisaf", "nisfNo");
});

app.get("/rukuh", (req, res) => {
  const surahNo = req.query.surahNo;
  const valueNo = req.query.rukuNo;

  // Load the Excel file
  const workbook = readFile("all-Distribuation.xlsx");

  // Get the specified sheet
  const sheet = workbook.Sheets["Ruku"];

  const filteredData = [];
  const range = sheet["!ref"].split(":");
  const startRow = parseInt(range[0].replace(/\D/g, ""), 10);
  const endRow = parseInt(range[1].replace(/\D/g, ""), 10);

  for (let i = startRow; i <= endRow; i++) {
    const surahValue = sheet[`A${i}`].v;

    if (surahValue == surahNo) {
      const rukuhNo = sheet[`B${i}`].v;
      const ayahNo = sheet[`C${i}`].v;
      filteredData.push({ rukuhNo, ayahNo });
    }
  }
  res.json(filteredData);
});

app.listen(3000, () => {
  console.log("Server started on port 3000");
});
