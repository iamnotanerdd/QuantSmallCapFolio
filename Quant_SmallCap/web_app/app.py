from flask import Flask, render_template, jsonify
import pandas as pd
import os
import math

app = Flask(__name__)

# Determine the base directory (root of the project)
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) # This is web_app/
ROOT_DIR = os.path.dirname(BASE_DIR) # Parent is Quant_SmallCap/
EXCEL_PATH = os.path.join(ROOT_DIR, "Quant_SmallCap_Equity_Analysis.xlsx")

def read_portfolio_data():
    if not os.path.exists(EXCEL_PATH):
        return {"error": "Excel file not found. Please run aggregator first."}

    # Structure matches PPFAS Aggregator output:
    # Row 1: Title
    # Row 2: Months (Merged)
    # Row 3: Subheaders
    # Row 4+: Data
    
    try:
        # Read Data Rows (Skip Title, Months, SubHeaders -> Start Row 4)
        df = pd.read_excel(EXCEL_PATH, header=None, skiprows=3) 
        
        # Read Month Row (Row 2, index 1)
        df_months = pd.read_excel(EXCEL_PATH, header=None, nrows=1, skiprows=1)
        month_row = df_months.iloc[0]
        
        # Extract unique ordered months from the merged cells row
        months = []
        # Columns 0,1,2 are static. Data starts at col 3.
        # Months are merged every 3 columns (3, 6, 9...)
        for i in range(3, len(month_row), 3):
            val = month_row[i]
            if pd.notna(val):
                months.append(str(val))
                
        data = []
        
        for _, row in df.iterrows():
            name = str(row[0])
            # Skip empty rows (except Total)
            if pd.isna(row[0]) or name.lower() == 'nan':
                continue
                
            def clean_meta(val):
                if pd.isna(val) or str(val).lower() == 'nan': return ""
                return str(val)

            record = {
                "Name": name, 
                "ISIN": clean_meta(row[1]),
                "Rating": clean_meta(row[2]),
                "Months": {}
            }
            
            col_idx = 3
            for m in months:
                def clean(val):
                    if pd.isna(val) or val == "": return 0.0
                    try: return float(val)
                    except: return 0.0
                    
                # Safe indexing
                if col_idx + 2 < len(row):
                    qty = clean(row[col_idx])
                    val = clean(row[col_idx+1])
                    pct = clean(row[col_idx+2])
                else:
                    qty, val, pct = 0, 0, 0
                
                record["Months"][m] = {
                    "Quantity": qty,
                    "Value": val,
                    "Pct": pct
                }
                col_idx += 3
                
            data.append(record)
            
        return {"months": months, "records": data}

    except Exception as e:
        print(f"Error reading Excel: {e}")
        return {"error": str(e)}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data')
def get_data():
    data = read_portfolio_data()
    return jsonify(data)

if __name__ == '__main__':
    print(f"Starting Web App...")
    print(f"Reading from: {EXCEL_PATH}")
    app.run(debug=True, port=5001) # Use 5001 to avoid conflict if PPFAS app is running
