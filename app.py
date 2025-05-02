from flask import Flask, jsonify, request, render_template
from flask_sqlalchemy import SQLAlchemy
import os
import pandas as pd
import numpy as np

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
    sheet = db.Column(db.String(50), nullable=False)

    def to_dict(self):
        return {
            'id': self.id,
            'row': self.row,
            'column': self.column,
            'value': self.value,
            'sheet': self.sheet
        }

def initialize_database_from_excel():
    try:
        sheets = ['inputs', 'outputs', 'outputs2']
        
        # Clear existing data
        Cell.query.delete()
        ColumnOrder.query.delete()
        
        for sheet_name in sheets:
            # Read the Excel file for each sheet
            df = pd.read_excel('small_book.xlsx', sheet_name=sheet_name)
            
            # Store column order
            for idx, column in enumerate(df.columns):
                col_order = ColumnOrder(column_name=column, order_index=idx, sheet=sheet_name)
                db.session.add(col_order)
            
            # Convert DataFrame to records
            for index, row in df.iterrows():
                for column in df.columns:
                    value = row[column]
                    if pd.notna(value):  # Only insert if value is not NaN
                        cell = Cell(
                            row=str(index + 1),  # Use row number as index
                            column=str(column),
                            value=str(value),
                            sheet=sheet_name
                        )
                        db.session.add(cell)
        
        db.session.commit()
        print("Database initialized from Excel file successfully")
    except Exception as e:
        print(f"Error initializing database: {str(e)}")
        db.session.rollback()

# Create all tables and initialize data
with app.app_context():
    db.create_all()
    initialize_database_from_excel()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/cells/<string:sheet>', methods=['GET'])
def get_cells(sheet):
    cells = Cell.query.filter_by(sheet=sheet).all()
    column_orders = ColumnOrder.query.filter_by(sheet=sheet).order_by(ColumnOrder.order_index).all()
    ordered_columns = [co.column_name for co in column_orders]
    
    return jsonify({
        'cells': [cell.to_dict() for cell in cells],
        'columnOrder': ordered_columns
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

if __name__ == '__main__':
    app.run(debug=True, port=5001) 