const axios = require('axios');
const ExcelJS = require('exceljs');
require('dotenv').config();

/**
 * Fetches information about a boat type from the Perplexity API
 * @param {string} boatType - The boat type to query for
 * @returns {Promise<string>} - The information retrieved from the API
 */
async function fetchBoatInfo(boatType) {
  try {
    // Get the query template from .env or use default
    const queryTemplate = process.env.QUERY_TEMPLATE || 
      "Provide detailed information about the boat type: {BOAT_TYPE}. Include details about its typical size, use cases, features, and history.";
    
    // Replace placeholder with actual boat type
    const query = queryTemplate.replace('{BOAT_TYPE}', boatType);
    
    // Make API call to Perplexity
    const response = await axios({
      method: 'post',
      url: 'https://api.perplexity.ai/chat/completions',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${process.env.PERPLEXITY_API_KEY}`
      },
      data: {
        model: 'mistral-7b-instruct',
        messages: [
          {
            role: 'user',
            content: query
          }
        ]
      }
    });
    
    // Extract information from the response
    const information = response.data.choices[0].message.content;
    return information;
  } catch (error) {
    console.error(`Error fetching info for ${boatType}:`, error.message);
    return `Error: Could not fetch information for ${boatType}`;
  }
}

/**
 * Reads boat types from an Excel file
 * @param {string} filePath - Path to the Excel file
 * @param {string} sheetName - Name of the sheet containing the data
 * @param {string} columnName - Name of the column containing boat types
 * @returns {Promise<Array<{row: number, boatType: string}>>} - Array of boat types with their row indices
 */
async function readBoatTypes(filePath, sheetName, columnName) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`Sheet ${sheetName} not found in the workbook`);
    }
    
    // Find the column index for the boat type
    let boatTypeColumnIndex = null;
    worksheet.getRow(1).eachCell((cell, colNumber) => {
      if (cell.value === columnName) {
        boatTypeColumnIndex = colNumber;
      }
    });
    
    if (!boatTypeColumnIndex) {
      throw new Error(`Column ${columnName} not found in the sheet`);
    }
    
    const boatTypes = [];
    
    // Start from row 2 (assuming row 1 is the header)
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const boatType = row.getCell(boatTypeColumnIndex).value;
        if (boatType) {
          boatTypes.push({ row: rowNumber, boatType: boatType.toString().trim() });
        }
      }
    });
    
    return boatTypes;
  } catch (error) {
    console.error('Error reading boat types from Excel:', error);
    throw error;
  }
}

/**
 * Updates the Excel file with boat information
 * @param {string} filePath - Path to the Excel file
 * @param {string} sheetName - Name of the sheet to update
 * @param {string} targetColumnName - Name of the column to update with boat info
 * @param {Array<{row: number, boatType: string, info: string}>} boatData - Array of boat data
 * @returns {Promise<void>}
 */
async function updateExcelWithInfo(filePath, sheetName, targetColumnName, boatData) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`Sheet ${sheetName} not found in the workbook`);
    }
    
    // Find or create the info column
    let infoColumnIndex = null;
    worksheet.getRow(1).eachCell((cell, colNumber) => {
      if (cell.value === targetColumnName) {
        infoColumnIndex = colNumber;
      }
    });
    
    if (!infoColumnIndex) {
      // If the column doesn't exist, create it
      infoColumnIndex = worksheet.columnCount + 1;
      worksheet.getCell(1, infoColumnIndex).value = targetColumnName;
    }
    
    // Update the data
    for (const boat of boatData) {
      worksheet.getCell(boat.row, infoColumnIndex).value = boat.info;
    }
    
    // Save the file
    await workbook.xlsx.writeFile(filePath);
    console.log(`Successfully updated ${boatData.length} rows in ${filePath}`);
  } catch (error) {
    console.error('Error updating Excel file:', error);
    throw error;
  }
}

/**
 * Creates a new Excel file with boat information if it doesn't exist
 * @param {string} filePath - Path to save the Excel file
 * @param {Array<string>} boatTypes - Array of boat types
 * @returns {Promise<void>}
 */
async function createExcelWithBoatTypes(filePath, boatTypes) {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Boat Types');
    
    // Add headers
    worksheet.columns = [
      { header: 'Boat Type', key: 'boatType', width: 20 },
      { header: 'Information', key: 'info', width: 80 }
    ];
    
    // Add rows for each boat type
    for (const boatType of boatTypes) {
      worksheet.addRow({ boatType, info: '' });
    }
    
    // Save the file
    await workbook.xlsx.writeFile(filePath);
    console.log(`Created new Excel file at ${filePath} with ${boatTypes.length} boat types`);
  } catch (error) {
    console.error('Error creating Excel file:', error);
    throw error;
  }
}

module.exports = {
  fetchBoatInfo,
  readBoatTypes,
  updateExcelWithInfo,
  createExcelWithBoatTypes
};
