# Boat Information Collector

A Node.js tool to automatically collect information about different boat types using the Perplexity API and save it to an Excel spreadsheet.

## Features

- Process an existing Excel file with boat types or create a new one
- Make API calls to Perplexity for detailed information about each boat type
- Customizable query template to get specific information
- Update the spreadsheet with retrieved information
- Rate limiting support to avoid API restrictions

## Prerequisites

- Node.js (v14 or higher)
- npm (comes with Node.js)
- A Perplexity API key

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/Tomski4200/boat-info-collector.git
   cd boat-info-collector
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. Create a `.env` file in the root directory and add your Perplexity API key:
   ```
   PERPLEXITY_API_KEY=your_api_key_here
   ```

## Usage

There are two main ways to use this tool:

### Option 1: Process an existing Excel file

If you already have an Excel file with a list of boat types, use this option:

```
node index.js --file yourfile.xlsx --sheet "Sheet1" --column "Boat Type" --target "Information"
```

Parameters:
- `--file` or `-f`: Path to your Excel file (required)
- `--sheet` or `-s`: Name of the sheet containing your data (default: "Sheet1")
- `--column` or `-c`: Column name containing boat types (default: "Boat Type")
- `--target` or `-t`: Column name where to store the information (default: "Information")
- `--delay` or `-d`: Delay between API calls in milliseconds (default: 1000)

### Option 2: Create a new Excel file with specified boat types

If you don't have an Excel file yet, you can create one with your list of boat types:

```
node index.js --create --list "Yacht,Sailboat,Catamaran,Dinghy" --output boats.xlsx
```

Parameters:
- `--create`: Flag to create a new file
- `--list` or `-l`: Comma-separated list of boat types
- `--output` or `-o`: Path for the output Excel file (default: "boat-types.xlsx")
- `--delay` or `-d`: Delay between API calls in milliseconds (default: 1000)

## Customizing API Queries

You can customize how the tool queries Perplexity by adding a `QUERY_TEMPLATE` to your `.env` file:

```
QUERY_TEMPLATE="Provide detailed information about the boat type: {BOAT_TYPE}. Include details about its typical size, use cases, features, and history."
```

The `{BOAT_TYPE}` placeholder will be replaced with each boat type.

## Examples

1. Process an existing file:
   ```
   node index.js -f data.xlsx -s "Boat Data" -c "Type" -t "Details" -d 2000
   ```

2. Create a new file with common boat types:
   ```
   node index.js --create --list "Yacht,Sailboat,Catamaran,Dinghy,Fishing Boat,Bowrider,Center Console,Pontoon Boat,Cabin Cruiser,Trawler,Houseboat" --output boat-types.xlsx
   ```

## Troubleshooting

- **API Key Issues**: Make sure your Perplexity API key is correctly set in the `.env` file.
- **Rate Limiting**: If you're getting rate limit errors, increase the delay between requests using the `--delay` option.
- **File Format**: Ensure your Excel file has headers in the first row.

## License

MIT
