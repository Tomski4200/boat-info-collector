#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
require('dotenv').config();

const {
  fetchBoatInfo,
  readBoatTypes,
  updateExcelWithInfo,
  createExcelWithBoatTypes
} = require('./utils');

// Define command line arguments
const argv = yargs(hideBin(process.argv))
  .option('file', {
    alias: 'f',
    description: 'Path to the Excel file',
    type: 'string'
  })
  .option('sheet', {
    alias: 's',
    description: 'Name of the sheet containing the boat types',
    type: 'string',
    default: 'Sheet1'
  })
  .option('column', {
    alias: 'c',
    description: 'Name of the column containing boat types',
    type: 'string',
    default: 'Boat Type'
  })
  .option('target', {
    alias: 't',
    description: 'Name of the column to write information to',
    type: 'string',
    default: 'Information'
  })
  .option('delay', {
    alias: 'd',
    description: 'Delay between API calls (in milliseconds)',
    type: 'number',
    default: 1000
  })
  .option('create', {
    description: 'Create a new file with the specified boat types',
    type: 'boolean',
    default: false
  })
  .option('list', {
    alias: 'l',
    description: 'Comma-separated list of boat types (for creating a new file)',
    type: 'string'
  })
  .option('output', {
    alias: 'o',
    description: 'Output file path for the new Excel file',
    type: 'string',
    default: 'boat-types.xlsx'
  })
  .help()
  .alias('help', 'h')
  .example('$0 -f data.xlsx -s "Boat Data" -c "Boat Type" -t "Details"', 'Process boat types from an existing file')
  .example('$0 --create --list "Yacht,Sailboat,Catamaran" -o boats.xlsx', 'Create a new file with specified boat types')
  .argv;

// Validate API key
if (!process.env.PERPLEXITY_API_KEY) {
  console.error('Error: Perplexity API key not found. Please set it in the .env file.');
  console.error('Create a .env file with PERPLEXITY_API_KEY=your_api_key_here');
  process.exit(1);
}

// Main function
async function main() {
  let boatTypesData = [];
  
  if (argv.create) {
    // Create a new file with boat types
    if (!argv.list) {
      console.error('Error: --list parameter is required when using --create');
      process.exit(1);
    }
    
    const boatTypes = argv.list.split(',').map(type => type.trim());
    
    console.log(`Creating a new Excel file with ${boatTypes.length} boat types...`);
    await createExcelWithBoatTypes(argv.output, boatTypes);
    
    // Prepare boat types for processing
    boatTypesData = boatTypes.map((boatType, index) => ({
      row: index + 2, // +2 because row 1 is header
      boatType
    }));
    
    // Update the file path to the newly created file
    argv.file = argv.output;
    argv.sheet = 'Boat Types';
  } else {
    // Process existing file
    if (!argv.file) {
      console.error('Error: --file parameter is required when processing an existing file');
      process.exit(1);
    }
    
    // Check if file exists
    if (!fs.existsSync(argv.file)) {
      console.error(`Error: File not found: ${argv.file}`);
      process.exit(1);
    }
    
    console.log(`Reading boat types from ${argv.file}, sheet ${argv.sheet}, column ${argv.column}...`);
    boatTypesData = await readBoatTypes(argv.file, argv.sheet, argv.column);
    console.log(`Found ${boatTypesData.length} boat types.`);
  }
  
  // Process each boat type
  console.log('Fetching information from Perplexity API...');
  const boatInfoData = [];
  
  for (let i = 0; i < boatTypesData.length; i++) {
    const { row, boatType } = boatTypesData[i];
    console.log(`[${i+1}/${boatTypesData.length}] Processing: ${boatType}`);
    
    // Fetch information from API
    const info = await fetchBoatInfo(boatType);
    
    boatInfoData.push({ row, boatType, info });
    
    // Add delay between API calls to avoid rate limiting
    if (i < boatTypesData.length - 1) {
      console.log(`Waiting ${argv.delay}ms before next request...`);
      await new Promise(resolve => setTimeout(resolve, argv.delay));
    }
  }
  
  // Update the Excel file with the retrieved information
  console.log(`Updating ${argv.file} with the retrieved information...`);
  await updateExcelWithInfo(argv.file, argv.sheet, argv.target, boatInfoData);
  
  console.log('Process completed successfully!');
}

// Run the main function
main()
  .catch(error => {
    console.error('An error occurred:', error);
    process.exit(1);
  });
