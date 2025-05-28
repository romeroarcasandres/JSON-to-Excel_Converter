# JSON to CSV/XLSX Converter
Advanced JSON to CSV and XLSX converter with interactive key selection and nested data support

## Overview:
This comprehensive script converts JSON files into both CSV and XLSX formats with advanced features for handling complex nested structures. It provides an interactive interface that allows users to select a JSON file, explore all available keys and subkeys, choose specific data fields for export, and generate clean tabular output files. The script intelligently handles nested objects, arrays, and mixed data structures while providing users full control over which data columns to include in the final export.

## Requirements:
- Python 3.6+
- tkinter library (for file dialog interface)
- pandas library (for data manipulation and XLSX export)
- openpyxl library (for Excel file creation)
- json library (for JSON parsing)
- csv library (for CSV handling)
- pathlib library (for file path operations)

## Files
JSON-to-Excel_Converter.py

## Installation
Before running the script, install the required dependencies:
```bash
pip install pandas openpyxl
```

## Usage
1. Run the script: `python JSON-to-Excel_Converter.py`
2. A file dialog will prompt you to select a JSON file
3. The script analyzes the JSON structure and displays all available leaf keys (nested endpoints)
4. Select the keys you want to export by entering numbers:
   - Individual keys: `1,3,5`
   - Ranges: `1-5,8,10-12`  
   - All keys: `all`
5. The script generates both CSV and XLSX files with the same base name as the input JSON file
6. Files are saved in the same directory as the source JSON file

## Key Features
- **Nested Structure Support**: Handles deeply nested JSON objects using dot notation (e.g., `user.profile.name`)
- **Array Processing**: Processes JSON arrays and creates separate rows for each array element
- **Smart Key Discovery**: Automatically discovers all keys across all objects, even when different objects have different key sets
- **Leaf Key Filtering**: Shows only meaningful endpoint keys, hiding redundant parent keys
- **Flexible Selection**: Interactive key selection with support for ranges and individual choices
- **Dual Format Export**: Generates both CSV and XLSX files simultaneously
- **Data Type Handling**: Intelligently converts complex nested objects to appropriate string representations
- **Missing Data Management**: Gracefully handles missing keys with null values
- **Unicode Support**: Full UTF-8 encoding support for international character sets

## Supported JSON Structures
- **Simple Objects**: `{"name": "John", "age": 30}`
- **Nested Objects**: `{"user": {"profile": {"name": "John", "city": "NYC"}}}`
- **Arrays of Objects**: `[{"name": "John"}, {"name": "Jane"}]`
- **Mixed Structures**: Complex combinations of nested objects and arrays
- **Variable Object Schemas**: Arrays where different objects have different key sets

## Example Usage
For a JSON file with structure:
```json
[
  {
    "source": {"language": "en", "value": "Hello"},
    "target": {"language": "es", "value": "Hola"}
  },
  {
    "source": {"language": "en", "value": "World"}, 
    "target": {"language": "es", "value": "Mundo"}
  }
]
```

The script will display:
```
Available keys for export:
----------------------------------------
 1. source.language
 2. source.value
 3. target.language
 4. target.value
```

And generate a CSV/XLSX with columns for each selected key and rows for each JSON array element.

## Important Notes
- Ensure the selected file is valid JSON format
- The script handles inconsistent key structures across different JSON objects
- Complex nested objects are converted to JSON string representation when necessary
- Array values are intelligently processed - single values extracted, complex arrays preserved as strings
- Output files use the same name as input JSON file with appropriate extensions (.csv, .xlsx)
- All file operations use UTF-8 encoding for international character support
- The script creates separate rows for each element in JSON arrays
- Missing keys in some objects result in null values in the corresponding CSV/XLSX cells

## Error Handling
- Invalid JSON files are detected and reported with clear error messages
- File access errors are handled gracefully with user-friendly notifications
- Missing dependencies are detected at startup with installation instructions
- Invalid key selections are caught with retry prompts

## License
This project is governed by the CC BY-NC 4.0 license. For comprehensive details, kindly refer to the LICENSE file included with this project.
