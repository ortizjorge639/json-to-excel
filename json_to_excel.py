#!/usr/bin/env python3
"""
JSON to Excel Converter for High/Low Order Text Data

This script processes JSON data containing high-order and low-order text relationships,
and generates an Excel file formatted according to a specific template.
"""

import json
import os
import argparse
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description='Convert JSON data to Excel format.')
    parser.add_argument('--input', '-i', required=True, help='Input JSON file path')
    parser.add_argument('--output', '-o', required=True, help='Output Excel file path')
    return parser.parse_args()

def load_json_data(file_path):
    """Load and validate JSON data from file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
        return data
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON format in {file_path}")
        exit(1)
    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
        exit(1)

def process_json_data(json_data):
    """
    Process JSON data to extract and format high-order and low-order text relationships.
    Extracts only from payload.results, ignores metadata, and ensures reasoning is shown only once per publication per high-order text.
    Returns a list of dictionaries ready for DataFrame creation.
    """
    rows = []

    # Support multiple JSON root structures
    if isinstance(json_data, dict) and 'payload' in json_data and 'results' in json_data['payload']:
        items = json_data['payload']['results']
    elif isinstance(json_data, list):
        items = json_data
    elif isinstance(json_data, dict):
        items = [json_data]
    else:
        print("Error: Unsupported JSON structure.")
        exit(1)

    for item in items:
        high_order_texts = item.get('high_order_text', [])
        reasonings = item.get('reasonings', [])
        # Build publication reasoning map for quick lookup
        pub_reasoning_map = {str(r.get('publication_ID', '')): r.get('reasoning', '') for r in reasonings}
        # For each high-order text
        for hot in high_order_texts:
            # Add high-order text row (no reasoning)
            high_order_row = {
                'Text Type': 'High-Order Text',
                'Paragraph ID': hot.get('paragraph_ID', ''),
                'Publication ID': hot.get('publication_ID', ''),
                'Task Text': hot.get('text', ''),
                'Tag': ', '.join(hot.get('tags', [])),
                'Similarity Score': 'N/A',
                'Reasonings': ''
            }
            rows.append(high_order_row)
            # Track which publications have shown reasoning for this high-order text
            reasoning_shown = set()
            # For each low-order text
            for lot in hot.get('low_order_texts', []):
                publication_id = str(lot.get('publication_ID', ''))
                paragraph_id = lot.get('paragraph_ID', '')
                # Only show reasoning for first occurrence of this publication_id under this high-order text
                reasoning_text = ''
                if publication_id and publication_id not in reasoning_shown:
                    reasoning_text = pub_reasoning_map.get(publication_id, '')
                    reasoning_shown.add(publication_id)
                low_order_row = {
                    'Text Type': 'Low-Order Text',
                    'Paragraph ID': paragraph_id,
                    'Publication ID': publication_id,
                    'Task Text': lot.get('text', ''),
                    'Tag': f"INCON-{hot.get('paragraph_ID', '')}",
                    'Similarity Score': lot.get('similarity_score', ''),
                    'Reasonings': reasoning_text
                }
                rows.append(low_order_row)
    return rows

def create_excel_file(data_rows, output_path):
    """
    Create an Excel file from the processed data.
    
    Args:
        data_rows: List of dictionaries containing row data
        output_path: Path for the output Excel file
    """
    # Create DataFrame from processed data
    df = pd.DataFrame(data_rows)
    
    # Write DataFrame to Excel without index
    df.to_excel(output_path, index=False, engine='openpyxl')
    
    # Apply formatting to match template
    format_excel_file(output_path)
    
    print(f"Excel file created successfully at: {output_path}")

def format_excel_file(file_path):
    """
    Apply formatting to the Excel file to match the template.
    
    Args:
        file_path: Path to the Excel file to format
    """
    # Load the workbook
    wb = load_workbook(file_path)
    ws = wb.active
    
    # Define styles
    header_font = Font(bold=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Format header row
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical='center')
        cell.border = border
    
    # Format data rows
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            cell.border = border
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        adjusted_width = min(max_length + 2, 50)  # Cap width at 50 characters
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Save the formatted workbook
    wb.save(file_path)

def main():
    """Main execution function."""
    # Parse command line arguments
    args = parse_arguments()
    
    # Load JSON data
    json_data = load_json_data(args.input)
    
    # Process JSON data
    processed_data = process_json_data(json_data)
    
    # Create Excel file
    create_excel_file(processed_data, args.output)

if __name__ == "__main__":
    main()