## 1. Project Overview

This Python script will process JSON data containing high-order and low-order text relationships, formatting this data into an Excel file that follows the structure of the provided template. No data modification or generation will occur during this process - the script will only extract and format the exact data from the JSON input.

## 2. Input Analysis

### 2.1 Excel Template Structure
The provided template has the following columns:
- Text Type
- Paragraph ID
- Publication ID
- Task Text
- Tag
- Similarity Score
- Reasonings

The template follows a pattern where:
- A high-order text is listed first
- Following rows contain related low-order texts
- Each high-order text and its associated low-order texts form a logical group

### 2.2 JSON Data Structure
The JSON data contains:
- `paragraph_id`: Unique identifier for the high-order paragraph
- `tag_type`: Identifies the tag as high-order ("hi_tag")
- `high_order_text`: Array containing high-order text information:
  - `text`: The actual high-order text content
  - `publication_ID`: Source publication identifier
  - `paragraph_ID`: Paragraph identifier within the publication
  - `tags`: Array of assigned tags
  - `low_order_texts`: Array of related low-order texts:
    - `text`: The low-order text content
    - `publication_ID`: Source publication identifier
    - `paragraph_ID`: Paragraph identifier
    - `tag`: Tag (null in the example)
    - `similarity_score`: Numeric score indicating similarity to high-order text
- `reasonings`: Array containing reasoning information for each publication ID

## 3. Functional Requirements

### 3.1 Data Processing
1. Parse the provided JSON data
2. Extract high-order text entries and their associated low-order texts exactly as provided
3. Format the data according to the Excel template structure without modifying any content
4. Ensure tags are preserved exactly as they appear in the JSON data

### 3.2 Excel Generation
1. Create an Excel workbook with a single sheet
2. Format the sheet with the required column headers
3. Populate rows with processed data following the template pattern:
   - High-order text row followed by its associated low-order text rows
4. Apply appropriate cell formatting to match the template

### 3.3 Reasoning Assignment
1. Include reasoning text for the first instance of each publication ID's low-order text
2. Leave reasoning cells empty for subsequent rows with the same publication ID

## 4. Technical Specifications

### 4.1 Required Libraries
- `pandas`: For data manipulation and Excel operations
- `openpyxl`: For Excel file formatting
- `json`: For parsing JSON data

### 4.2 Input Requirements
- JSON data file with the specified structure
- Location of the input file should be configurable

### 4.3 Output Requirements
- Excel file (.xlsx) with formatting matching the template
- Output file name and location should be configurable

### 4.4 Error Handling
- Validate JSON format before processing
- Handle missing or null values in the data
- Provide informative error messages for common issues

## 5. Implementation Plan

### 5.1 Data Transformation Logic
1. Read JSON data into memory
2. Create a data structure to hold transformed data
3. Process each high-order text entry:
   - Extract high-order text information exactly as provided in the JSON
   - Extract related low-order text information exactly as provided in the JSON
   - Use existing INCON relationships from the data
   - Associate reasoning data with appropriate entries (once per publication ID)
4. Assemble the data into a format suitable for Excel export without modifying any information

### 5.2 Excel Output Generation
1. Create a pandas DataFrame from the extracted data
2. Apply formatting to match the template
3. Write the formatted DataFrame to an Excel file
4. Add any additional formatting required to match the template

## 6. Usage Instructions

1. Install the required Python packages:
   ```
   pip install pandas openpyxl
   ```

2. Run the script with the appropriate arguments:
   ```
   python json_to_excel_converter.py --input input_data.json --output formatted_data.xlsx
   ```

   Where:
   - `input_data.json` is the path to your JSON input file
   - `formatted_data.xlsx` is the desired path for the Excel output file

## 7. Key Features

1. **Faithful Data Representation**: 
   - Preserves all original data without modifications
   - Maintains the hierarchical relationship between high-order and low-order texts

2. **Proper Formatting**:
   - Formats Excel output to match the template
   - Applies appropriate cell formatting for readability

3. **Reasoning Assignment**:
   - Associates reasoning text with the appropriate low-order text entries
   - Ensures reasoning is only displayed once per publication ID/paragraph ID combination

4. **Error Handling**:
   - Validates input data format
   - Provides clear error messages for troubleshooting

## 8. Testing Plan

1. **Unit Testing**:
   - Test JSON parsing with valid and invalid inputs
   - Verify correct data extraction without modifications
   - Confirm proper Excel file generation

2. **Integration Testing**:
   - End-to-end test with sample JSON data
   - Validate output against expected Excel format

3. **Edge Cases**:
   - Handle empty JSON data
   - Process JSON with missing fields
   - Manage large datasets efficiently

## 9. Special Considerations

1. The script maintains data integrity by:
   - Never generating or modifying any data values
   - Preserving all original relationships and tags as provided in the JSON
   - Only formatting the data to match the required Excel structure

2. The reasoning text is applied only once per unique publication ID/paragraph ID combination, following the pattern seen in the template example.

CRITICAL RULES:
Try to fix things at the cause, not the symptom.

Be very detailed with summarization and do not miss out things that are important.