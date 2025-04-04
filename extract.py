import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import sys
from collections import defaultdict

def extract_tables(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Attempt built-in table extraction
            page_tables = page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 5,
                "join_tolerance": 5,
                "edge_min_length": 15,
                "min_words_vertical": 2,
                "min_words_horizontal": 1
            })
            if page_tables:
                tables.extend(page_tables)
            else:
                # Fallback to custom extraction
                words = page.extract_words(keep_blank_chars=False, extra_attrs=["fontname", "size"])
                if words:
                    custom_table = custom_extract_table(words)
                    if custom_table:
                        tables.append(custom_table)
    return tables

def custom_extract_table(words):
    # Group words into rows based on y-coordinate
    rows = defaultdict(list)
    tolerance = 5  # Adjust based on testing
    for word in words:
        y = word['top']
        matched = False
        for key in list(rows.keys()):
            if abs(y - key) <= tolerance:
                rows[key].append(word)
                matched = True
                break
        if not matched:
            rows[y].append(word)
    
    # Sort rows by y-coordinate
    sorted_rows = sorted(rows.items(), key=lambda x: x[0])
    table_rows = [sorted(row[1], key=lambda w: w['x0']) for row in sorted_rows]
    
    # Determine column boundaries
    x_positions = []
    for row in table_rows:
        for word in row:
            x_positions.append((word['x0'], word['x1']))
    
    # Find common column boundaries (simplified)
    columns = []
    if x_positions:
        x_positions.sort()
        current_start, current_end = x_positions[0]
        for start, end in x_positions[1:]:
            if start <= current_end + tolerance:
                current_end = max(current_end, end)
            else:
                columns.append((current_start, current_end))
                current_start, current_end = start, end
        columns.append((current_start, current_end))
    
    # Assign words to columns
    table_data = []
    for row in table_rows:
        row_data = [''] * len(columns)
        for word in row:
            mid_x = (word['x0'] + word['x1']) / 2
            for idx, (start, end) in enumerate(columns):
                if start <= mid_x <= end:
                    if row_data[idx] == '':
                        row_data[idx] = word['text']
                    else:
                        row_data[idx] += ' ' + word['text']
                    break
        table_data.append(row_data)
    return table_data

def save_to_excel(tables, output_path):
    wb = Workbook()
    wb.remove(wb.active)
    for i, table in enumerate(tables):
        ws = wb.create_sheet(title=f"Table_{i+1}")
        for row in table:
            ws.append(row)
    wb.save(output_path)

def process_pdf(pdf_path, output_dir):
    tables = extract_tables(pdf_path)
    if not tables:
        print(f"No tables found in {pdf_path}")
        return
    output_file = os.path.join(output_dir, os.path.splitext(os.path.basename(pdf_path))[0] + '.xlsx')
    save_to_excel(tables, output_file)
    print(f"Saved tables to {output_file}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python script.py <input_directory> <output_directory>")
        sys.exit(1)
    input_dir, output_dir = sys.argv[1], sys.argv[2]
    os.makedirs(output_dir, exist_ok=True)
    for filename in os.listdir(input_dir):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(input_dir, filename)
            process_pdf(pdf_path, output_dir)