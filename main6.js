const fs = require("fs");
const xlsx = require("xlsx");

// Function to read Excel file and return data as JSON
const readExcelFile = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(worksheet);
};

// Function to read JSON file
const readJsonFile = (filePath) => {
  const data = fs.readFileSync(filePath, "utf8");
  return JSON.parse(data);
};

// Function to normalize data (trim strings and convert to proper types)
const normalizeValue = (value) => {
  const isValidDate = (date) => !isNaN(new Date(date).getTime());
  const isNumeric = (value) => !isNaN(parseFloat(value)) && isFinite(value);

  if (typeof value === "string") {
    const trimmedValue = value.trim();

    if (isValidDate(trimmedValue)) {
      return new Date(trimmedValue).toISOString().split("T")[0];
    }

    if (isNumeric(trimmedValue)) {
      return parseFloat(trimmedValue);
    }

    return trimmedValue;
  }

  if (typeof value === "number") {
    // if (value > 0 && value < 60000) {
    //   const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    //   const dateValue = new Date(excelEpoch.getTime() + value * 86400000);
    //   return dateValue.toISOString().split("T")[0];
    // }

    return value;
  }

  return value;
};

// Function to compare JSON and Excel data
const compareData = (jsonData, excelData) => {
  return excelData.reduce((differences, excelRow, index) => {
    const jsonRow = jsonData[index];
    const rowDiff = {};

    Object.keys(excelRow).forEach((key) => {
      const excelValue = normalizeValue(excelRow[key]);
      const jsonValue = jsonRow ? normalizeValue(jsonRow[key]) : undefined;

      if (excelValue !== jsonValue) {
        rowDiff[key] = {
          excelValue,
          jsonValue,
        };
      }
    });

    if (Object.keys(rowDiff).length > 0) {
      differences.push({ rowIndex: index, differences: rowDiff });
    }
    return differences;
  }, []);
};

// Main function
const main = () => {
  const jsonFilePath = "C:/Users/Unique/Downloads/Extracted_Text.json"; // Replace with the path to your JSON file
  const excelFilePath = "C:/Users/Unique/Downloads/convert.xlsx"; // Replace with the path to your Excel file

  try {
    const jsonData = readJsonFile(jsonFilePath);
    const excelData = readExcelFile(excelFilePath);

    // Compare data and get differences
    const differences = compareData(jsonData, excelData);

    console.log("Differences:", JSON.stringify(differences, null, 2));
  } catch (error) {
    console.error("Error reading files:", error);
  }
};

// Execute main function
main();
