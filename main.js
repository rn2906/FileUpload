const axios = require("axios");
const xlsx = require("xlsx");

// Function to read Excel file and return data as JSON
const readExcelFile = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(worksheet);
};

// Function to compare JSON and Excel data
const compareData = (jsonData, excelData) => {
  return excelData.reduce((differences, excelRow, index) => {
    const jsonRow = jsonData[index];
    const rowDiff = Object.keys(excelRow).reduce((diff, key) => {
      if (excelRow[key] !== (jsonRow ? jsonRow[key] : undefined)) {
        diff[key] = {
          excelValue: excelRow[key],
          jsonValue: jsonRow ? jsonRow[key] : null,
        };
      }
      return diff;
    }, {});

    if (Object.keys(rowDiff).length > 0) {
      differences.push({ rowIndex: index, differences: rowDiff });
    }
    return differences;
  }, []);
};

// Main function
const main = async () => {
  const apiUrl = "https://api.example.com/data"; // Replace with your API URL
  const excelFilePath = "path/to/your/excelfile.xlsx"; // Replace with the path to your Excel file

  try {
    const { data: jsonData } = await axios.get(apiUrl);
    const excelData = readExcelFile(excelFilePath);

    const differences = compareData(jsonData, excelData);

    console.log("Differences:", JSON.stringify(differences, null, 2));
  } catch (error) {
    console.error("Error fetching JSON data:", error);
  }
};

// Execute main function
main();
