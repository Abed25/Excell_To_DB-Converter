const express = require("express");
const XLSX = require("xlsx");
const mysql = require("mysql");

const app = express();
const port = 3000;

// MySQL connection setup
const connection = mysql.createConnection({
  host: "127.0.0.1",
  user: "root",
  password: "",
  database: "epza-info",
  port: 3308, // Specify your MySQL port here
});

connection.connect((err) => {
  if (err) {
    console.error("Error connecting to MySQL: " + err.stack);
    return;
  }
  console.log("Connected to MySQL as id " + connection.threadId);
});

// Function to remove commas from numerical strings and convert to integers
const parseIntWithCommas = (value) => {
  if (typeof value === "string") {
    // If the value is zero but formatted as string, convert to 0
    if (value.trim() === "0") {
      return 0;
    }
    return parseInt(value.replace(/,/g, ""), 10);
  }
  return value;
};

// Express route to handle file upload and conversion
app.post("/upload", (req, res) => {
  const excelFilePath = "C:/Users/Abed/Documents/test2.xlsx"; // Specify your Excel file path here

  try {
    const workbook = XLSX.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    // Assuming the first row contains headers
    const headers = data[0];
    const records = data.slice(1);

    // Define the columns that should be integers
    const intColumns = ["domestic_sale", "daleDF", "exports", "total_sales"];

    // Ensuring all records have the same number of columns
    const completeRecords = records.map((record) => {
      return headers.map((header, index) => {
        let value = record[index];
        // Remove commas from numerical values and convert to integers, set to 0 if empty or zero
        if (intColumns.includes(header)) {
          if (value === "" || value === null) {
            return 0;
          }
          return parseIntWithCommas(value);
        }
        return value || null;
      });
    });

    // Wrap headers with backticks
    const wrappedHeaders = headers.map((header) => `\`${header}\``);

    // Constructing SQL query to insert or update data into MySQL
    let sql = `INSERT INTO test2 (${wrappedHeaders.join(", ")}) VALUES ? 
                   ON DUPLICATE KEY UPDATE `;

    // Append update statements for each column
    sql += headers
      .map((header) => `\`${header}\` = VALUES(\`${header}\`)`)
      .join(", ");

    connection.query(sql, [completeRecords], (err, result) => {
      if (err) {
        console.error(
          "Error inserting or updating data into MySQL: " + err.message
        );
        res.status(500).send("Error inserting or updating data into MySQL");
        return;
      }
      console.log(
        "Number of records inserted or updated: " + result.affectedRows
      );
      res.send("Data imported and updated successfully");
    });
  } catch (error) {
    console.error("Error processing Excel file: " + error.message);
    res.status(500).send("Error processing Excel file");
  }
});

// Start server
app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});
