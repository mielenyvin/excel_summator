from flask import Flask, jsonify, request, render_template
from flask_sqlalchemy import SQLAlchemy
import os
import pandas as pd
import numpy as np
import re
import openpyxl
from functools import lru_cache
from collections import defaultdict

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///grid.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# Delete the database file if it exists
if os.path.exists('grid.db'):
    os.remove('grid.db')

class ColumnOrder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    column_name = db.Column(db.String(50), nullable=False)
    order_index = db.Column(db.Integer, nullable=False)
    sheet = db.Column(db.String(50), nullable=False)

    __table_args__ = (
        db.UniqueConstraint('column_name', 'sheet', name='uix_column_sheet'),
    )

class Cell(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    row = db.Column(db.String(50), nullable=False)
    column = db.Column(db.String(50), nullable=False)
    value = db.Column(db.String(50), nullable=True)
    formula = db.Column(db.String(50), nullable=True)  # Store original formula
    sheet = db.Column(db.String(50), nullable=False)
    excel_row = db.Column(db.String(50), nullable=True)
    excel_col = db.Column(db.String(50), nullable=True)
    excel_coord = db.Column(db.String(50), nullable=True)

    def to_dict(self):
        return {
            'id': self.id,
            'row': self.row,
            'column': self.column,
            'value': self.value,
            'formula': self.formula,  # Include formula in response
            'sheet': self.sheet,
            'excel_row': self.excel_row,
            'excel_col': self.excel_col,
            'excel_coord': self.excel_coord
        }

def get_column_letter(column_index):
    """Convert column index to Excel-style letter (1->A, 2->B, etc.)"""
    result = ""
    while column_index > 0:
        column_index, remainder = divmod(column_index - 1, 26)
        result = chr(65 + remainder) + result
    return result

def initialize_database_from_excel():
    try:
        sheets = ['inputs', 'outputs', 'outputs2']
        
        # Clear existing data
        Cell.query.delete()
        ColumnOrder.query.delete()
        
        for sheet_name in sheets:
            print(f"\nProcessing sheet: {sheet_name}")
            # Read the Excel file for each sheet using openpyxl to preserve formulas
            wb = openpyxl.load_workbook('small_book.xlsx', data_only=False)
            ws = wb[sheet_name]
            
            # Get column names from first row
            columns = []
            for cell in ws[1]:
                columns.append(cell.value)
            print(f"Columns in sheet: {columns}")
            
            # Store column order
            for idx, column in enumerate(columns):
                col_order = ColumnOrder(column_name=column, order_index=idx, sheet=sheet_name)
                db.session.add(col_order)
            
            # Process each cell
            for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=1):  # Start from row 2 (after header)
                print(f"\nProcessing row {row_idx}:")
                for col_idx, cell in enumerate(row):
                    column = columns[col_idx]
                    if cell.value is not None:  # Only process non-empty cells
                        value = cell.value
                        print(f"  Cell: column={column}, value={value}, type={type(value)}")
                        
                        # Store Excel coordinates
                        excel_col = get_column_letter(col_idx + 1)  # +1 because Excel is 1-based
                        excel_row = row_idx + 1  # +1 because we started from row 2
                        excel_coord = f"{excel_col}{excel_row}"
                        
                        # Check if it's a formula
                        if sheet_name != 'inputs' and cell.data_type == 'f':  # 'f' means formula
                            # Ensure formula starts with =
                            str_value = str(value)
                            if not str_value.startswith('='):
                                str_value = '=' + str_value
                            print(f"  Found formula: {str_value}")
                            formula = str_value
                            value = None  # Initial value is None for formula cells
                        else:
                            str_value = str(value)
                            formula = None
                        
                        cell_db = Cell(
                            row=str(row_idx),
                            column=str(column),
                            value=str_value,
                            formula=formula,
                            sheet=sheet_name,
                            excel_row=str(excel_row),
                            excel_col=excel_col,
                            excel_coord=excel_coord
                        )
                        db.session.add(cell_db)
            
            # Commit after each sheet to see progress
            db.session.commit()
            print(f"Completed processing sheet: {sheet_name}")
        
        print("Database initialized from Excel file successfully")
    except Exception as e:
        print(f"Error initializing database: {str(e)}")
        db.session.rollback()

# Create all tables and initialize data
with app.app_context():
    db.create_all()
    initialize_database_from_excel()

# Cache for column orders
@lru_cache(maxsize=32)
def get_column_order(sheet):
    return ColumnOrder.query.filter_by(sheet=sheet).order_by(ColumnOrder.order_index).all()

# Cache for lookup arrays
@lru_cache(maxsize=32)
def get_lookup_array(sheet, column):
    """Get all values from a column in a sheet, cached"""
    cells = Cell.query.filter_by(sheet=sheet, excel_col=column).order_by(Cell.excel_row).all()
    return [(cell.excel_row, cell.value) for cell in cells]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/cells/<string:sheet>', methods=['GET'])
def get_cells(sheet):
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 100, type=int)
    
    # Get total count of rows
    total_rows = db.session.query(db.func.count(db.distinct(Cell.row))).filter_by(sheet=sheet).scalar()
    total_pages = (total_rows + per_page - 1) // per_page
    
    # Get unique row numbers for the current page
    row_numbers = db.session.query(Cell.row).filter_by(sheet=sheet).distinct().order_by(Cell.row).offset((page - 1) * per_page).limit(per_page).all()
    row_numbers = [row[0] for row in row_numbers]
    
    # Get column order from cache
    column_orders = get_column_order(sheet)
    ordered_columns = [co.column_name for co in column_orders]
    
    # Get only cells for these rows and columns
    cells = Cell.query.filter_by(sheet=sheet).filter(
        Cell.row.in_(row_numbers),
        Cell.column.in_(ordered_columns)
    ).all()
    
    return jsonify({
        'cells': [cell.to_dict() for cell in cells],
        'columnOrder': ordered_columns,
        'pagination': {
            'current_page': page,
            'total_pages': total_pages,
            'total_rows': total_rows,
            'per_page': per_page
        }
    })

@app.route('/api/cells/<string:sheet>/<int:row>/<string:column>', methods=['PUT'])
def update_cell(sheet, row, column):
    cell = Cell.query.filter_by(sheet=sheet, row=row, column=column).first()
    if cell:
        data = request.get_json()
        cell.value = data.get('value', cell.value)
        db.session.commit()
        return jsonify(cell.to_dict())
    return jsonify({'error': 'Cell not found'}), 404

@app.route('/api/cells/<string:sheet>/save-all', methods=['POST'])
def save_all_cells(sheet):
    try:
        updates = request.get_json()
        for update in updates:
            cell = Cell.query.filter_by(sheet=sheet, row=update['row'], column=update['column']).first()
            if cell:
                # Handle empty string and zero values properly
                value = update['value']
                if value == '':
                    cell.value = None
                else:
                    cell.value = value
        db.session.commit()
        return jsonify({'status': 'success'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

def get_cell_value(sheet, row, column):
    cell = Cell.query.filter_by(sheet=sheet, row=str(row), column=column).first()
    if cell and cell.value:
        try:
            # If the value starts with =, it's a formula, so we need to evaluate it
            if cell.value.startswith('='):
                return evaluate_formula(cell.value[1:], sheet)
            return float(cell.value)
        except ValueError:
            return cell.value
    return 0

def xlookup(lookup_value, lookup_array, return_array, if_not_found="Error"):
    try:
        # Convert lookup_value to string for comparison
        lookup_value = str(lookup_value)
        
        # Find the index of the lookup value in the lookup array
        for i, value in enumerate(lookup_array):
            if str(value) == lookup_value:
                return return_array[i]
        
        return if_not_found
    except Exception as e:
        print(f"Error in XLOOKUP: {str(e)}")
        return "Error"

def evaluate_formula(formula, sheet_name):
    try:
        print(f"\nEvaluating formula: {formula} in sheet: {sheet_name}")
        
        # Remove the leading '=' if present
        formula = formula.lstrip('=')
        print(f"Formula after stripping '=': {formula}")
        
        # Handle XLOOKUP formula
        if '_xlfn.XLOOKUP' in formula:
            print("Detected XLOOKUP formula")
            # Extract parameters from XLOOKUP formula
            params_str = formula.replace('_xlfn.XLOOKUP', '').strip('()')
            params = []
            current_param = ""
            paren_count = 0
            for char in params_str:
                if char == '(':
                    paren_count += 1
                elif char == ')':
                    paren_count -= 1
                elif char == ',' and paren_count == 0:
                    params.append(current_param.strip())
                    current_param = ""
                    continue
                current_param += char
            params.append(current_param.strip())
            
            print(f"XLOOKUP parameters: {params}")
            
            if len(params) < 3:
                print("Invalid XLOOKUP formula - not enough parameters")
                return "Invalid XLOOKUP formula"
            
            lookup_value = params[0].strip()
            lookup_array = params[1].strip()
            return_array = params[2].strip()
            
            print(f"Lookup value: {lookup_value}")
            print(f"Lookup array: {lookup_array}")
            print(f"Return array: {return_array}")
            
            # Extract Excel coordinates from lookup value (e.g., A2)
            lookup_coord = lookup_value
            print(f"Lookup coordinate: {lookup_coord}")
            
            # Get the value to look up
            lookup_cell = Cell.query.filter_by(
                excel_coord=lookup_coord,
                sheet=sheet_name
            ).first()
            
            if not lookup_cell:
                print(f"Lookup cell not found: coord={lookup_coord}")
                return "Lookup value not found"
            
            lookup_value = lookup_cell.value
            print(f"Value to look up: {lookup_value}")
            
            # Get lookup and return arrays from cache
            lookup_sheet = lookup_array.split('!')[0]
            lookup_col = lookup_array.split('!')[1].split(':')[0]
            lookup_cells = get_lookup_array(lookup_sheet, lookup_col)
            
            # Find the matching value in lookup array
            match_index = -1
            for i, (row, value) in enumerate(lookup_cells):
                if value == lookup_value:
                    match_index = i
                    break
            
            if match_index == -1:
                print(f"No match found for value: {lookup_value}")
                return "No match found"
            
            # Get return array from cache
            return_sheet = return_array.split('!')[0]
            return_col = return_array.split('!')[1].split(':')[0]
            return_cells = get_lookup_array(return_sheet, return_col)
            
            if match_index >= len(return_cells):
                print(f"Return array index out of bounds: {match_index}")
                return "Index out of bounds"
            
            result = return_cells[match_index][1]  # Get value from (row, value) tuple
            print(f"Found matching value: {result}")
            return result
        
        # Handle SUM formula
        elif formula.startswith('SUM('):
            print("Detected SUM formula")
            # Extract range from SUM formula
            range_str = formula[4:].strip('()')
            print(f"SUM range: {range_str}")
            
            # Extract sheet and range
            sheet, range_cells = range_str.split('!')
            print(f"Sheet: {sheet}, Range: {range_cells}")
            
            # Extract start and end cells
            start_cell, end_cell = range_cells.split(':')
            print(f"Start cell: {start_cell}, End cell: {end_cell}")
            
            # Extract column and row numbers
            start_col = ''.join(filter(str.isalpha, start_cell))
            start_row = int(''.join(filter(str.isdigit, start_cell)))
            end_col = ''.join(filter(str.isalpha, end_cell))
            end_row = int(''.join(filter(str.isdigit, end_cell)))
            
            # Get all cells in the range from cache
            cells = get_lookup_array(sheet, start_col)
            sum_value = 0
            
            for row, value in cells:
                row_num = int(row)
                if start_row <= row_num <= end_row:
                    try:
                        sum_value += float(value)
                    except (ValueError, TypeError):
                        continue
            
            print(f"SUM result: {sum_value}")
            return str(sum_value)
        
        # Handle regular formulas
        try:
            print("Evaluating regular formula")
            
            # Handle concatenation formulas (e.g., =A2&B2)
            if '&' in formula:
                print("Detected concatenation formula")
                # Split by & and get cell references
                cell_refs = formula.split('&')
                result = ""
                
                for ref in cell_refs:
                    # Extract column and row from reference (e.g., A2)
                    col = ''.join(filter(str.isalpha, ref))
                    row = int(''.join(filter(str.isdigit, ref)))
                    
                    # Get the cell value
                    cell = Cell.query.filter_by(
                        sheet=sheet_name,
                        excel_col=col,
                        excel_row=str(row)
                    ).first()
                    
                    if cell and cell.value:
                        result += str(cell.value)
                    else:
                        result += ""
                
                print(f"Concatenation result: {result}")
                return result
            
            # Handle other regular formulas
            result = eval(formula)
            print(f"Formula result: {result}")
            return str(result)
        except Exception as e:
            print(f"Error evaluating regular formula: {str(e)}")
            return f"Error evaluating formula: {str(e)}"
    except Exception as e:
        print(f"General error in evaluate_formula: {str(e)}")
        return f"Error: {str(e)}"

def calculate_sheet(sheet):
    if sheet == 'outputs':
        source_sheet = 'inputs'
    elif sheet == 'outputs2':
        source_sheet = 'outputs'
    else:
        return

    # Get all cells for the current sheet
    cells = Cell.query.filter_by(sheet=sheet).all()
    
    for cell in cells:
        if cell.value and cell.value.startswith('='):
            # This is a formula cell
            formula = cell.value[1:]  # Remove the = sign
            result = evaluate_formula(formula, source_sheet)
            cell.value = result
    
    db.session.commit()

@app.route('/calculate', methods=['POST'])
def calculate():
    try:
        # Get all cells from the current sheet
        current_sheet = request.json.get('sheet', 'inputs')
        print(f"\nCalculating formulas for sheet: {current_sheet}")
        
        cells = Cell.query.filter_by(sheet=current_sheet).all()
        print(f"Found {len(cells)} cells in sheet {current_sheet}")
        
        # Print all cells for debugging
        print("\nAll cells in sheet:")
        for cell in cells:
            print(f"Row: {cell.row}, Column: {cell.column}, Value: {cell.value}, Formula: {cell.formula}")
        
        # Process each cell
        for cell in cells:
            if cell.formula:  # If cell has a formula
                print(f"\nProcessing formula cell: row={cell.row}, column={cell.column}")
                print(f"Original formula: {cell.formula}")
                
                # Evaluate the formula
                result = evaluate_formula(cell.formula, current_sheet)
                print(f"Evaluation result: {result}")
                
                # Update value but keep formula
                cell.value = result
                print(f"Updated cell value: {cell.value}")
            else:
                print(f"\nNon-formula cell: row={cell.row}, column={cell.column}, value={cell.value}")
        
        db.session.commit()
        print("\nCalculation completed successfully")
        return jsonify({'status': 'success'})
    except Exception as e:
        print(f"\nError in calculate: {str(e)}")
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)})

if __name__ == '__main__':
    app.run(debug=True, port=5001) 