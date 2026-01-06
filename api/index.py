from flask import Flask, render_template, request, jsonify
import openpyxl
import os
import platform
import subprocess
from datetime import datetime

app = Flask(__name__)

# FULL DATASET FROM ALL 8 RAPAPORT TABLES
RAP_DATA = [
    {"min": 0.30, "max": 0.39, "data": {"D": {"IF": 31, "VVS1": 24, "VVS2": 21, "VS1": 19, "VS2": 17, "SI1": 15, "SI2": 14, "SI3": 13, "I1": 12, "I2": 11, "I3": 7}, "E": {"IF": 26, "VVS1": 22, "VVS2": 19, "VS1": 17, "VS2": 16, "SI1": 14, "SI2": 13, "SI3": 12, "I1": 11, "I2": 10, "I3": 6}, "F": {"IF": 23, "VVS1": 20, "VVS2": 18, "VS1": 16, "VS2": 15, "SI1": 13, "SI2": 12, "SI3": 11, "I1": 11, "I2": 10, "I3": 6}, "G": {"IF": 20, "VVS1": 18, "VVS2": 16, "VS1": 15, "VS2": 14, "SI1": 13, "SI2": 12, "SI3": 11, "I1": 10, "I2": 9, "I3": 5}, "H": {"IF": 17, "VVS1": 16, "VVS2": 15, "VS1": 14, "VS2": 13, "SI1": 12, "SI2": 11, "SI3": 10, "I1": 9, "I2": 8, "I3": 5}, "I": {"IF": 15, "VVS1": 14, "VVS2": 13, "VS1": 12, "VS2": 11, "SI1": 11, "SI2": 10, "SI3": 9, "I1": 8, "I2": 7, "I3": 5}, "J": {"IF": 13, "VVS1": 12, "VVS2": 11, "VS1": 11, "VS2": 10, "SI1": 10, "SI2": 10, "SI3": 9, "I1": 8, "I2": 7, "I3": 4}, "K": {"IF": 12, "VVS1": 11, "VVS2": 10, "VS1": 9, "VS2": 9, "SI1": 9, "SI2": 9, "SI3": 8, "I1": 7, "I2": 6, "I3": 4}, "L": {"IF": 11, "VVS1": 10, "VVS2": 9, "VS1": 8, "VS2": 8, "SI1": 8, "SI2": 8, "SI3": 7, "I1": 6, "I2": 5, "I3": 3}, "M": {"IF": 10, "VVS1": 9, "VVS2": 9, "VS1": 8, "VS2": 8, "SI1": 8, "SI2": 7, "SI3": 6, "I1": 5, "I2": 4, "I3": 3}}},
    {"min": 0.40, "max": 0.49, "data": {"D": {"IF": 35, "VVS1": 28, "VVS2": 24, "VS1": 22, "VS2": 20, "SI1": 18, "SI2": 16, "SI3": 15, "I1": 14, "I2": 12, "I3": 8}, "E": {"IF": 29, "VVS1": 25, "VVS2": 22, "VS1": 20, "VS2": 19, "SI1": 17, "SI2": 15, "SI3": 14, "I1": 13, "I2": 11, "I3": 7}, "F": {"IF": 26, "VVS1": 23, "VVS2": 21, "VS1": 19, "VS2": 18, "SI1": 16, "SI2": 14, "SI3": 13, "I1": 12, "I2": 11, "I3": 7}, "G": {"IF": 23, "VVS1": 20, "VVS2": 19, "VS1": 18, "VS2": 17, "SI1": 15, "SI2": 13, "SI3": 12, "I1": 11, "I2": 10, "I3": 6}, "H": {"IF": 20, "VVS1": 18, "VVS2": 17, "VS1": 16, "VS2": 15, "SI1": 14, "SI2": 12, "SI3": 11, "I1": 10, "I2": 9, "I3": 6}, "I": {"IF": 18, "VVS1": 16, "VVS2": 15, "VS1": 14, "VS2": 13, "SI1": 12, "SI2": 11, "SI3": 10, "I1": 9, "I2": 8, "I3": 6}, "J": {"IF": 16, "VVS1": 14, "VVS2": 13, "VS1": 13, "VS2": 12, "SI1": 11, "SI2": 11, "SI3": 10, "I1": 9, "I2": 8, "I3": 5}, "K": {"IF": 14, "VVS1": 13, "VVS2": 12, "VS1": 11, "VS2": 11, "SI1": 10, "SI2": 10, "SI3": 9, "I1": 8, "I2": 7, "I3": 5}, "L": {"IF": 13, "VVS1": 12, "VVS2": 11, "VS1": 10, "VS2": 10, "SI1": 9, "SI2": 9, "SI3": 8, "I1": 7, "I2": 6, "I3": 4}, "M": {"IF": 12, "VVS1": 11, "VVS2": 10, "VS1": 9, "VS2": 9, "SI1": 9, "SI2": 8, "SI3": 7, "I1": 6, "I2": 5, "I3": 4}}},
    {"min": 0.50, "max": 0.69, "data": {"D": {"IF": 55, "VVS1": 44, "VVS2": 34, "VS1": 28, "VS2": 25, "SI1": 22, "SI2": 18, "SI3": 16, "I1": 15, "I2": 14, "I3": 11}, "E": {"IF": 44, "VVS1": 38, "VVS2": 31, "VS1": 26, "VS2": 23, "SI1": 20, "SI2": 17, "SI3": 15, "I1": 14, "I2": 13, "I3": 10}, "F": {"IF": 38, "VVS1": 33, "VVS2": 28, "VS1": 24, "VS2": 22, "SI1": 19, "SI2": 16, "SI3": 14, "I1": 13, "I2": 12, "I3": 10}, "G": {"IF": 32, "VVS1": 28, "VVS2": 25, "VS1": 23, "VS2": 21, "SI1": 18, "SI2": 15, "SI3": 13, "I1": 12, "I2": 11, "I3": 9}, "H": {"IF": 26, "VVS1": 23, "VVS2": 22, "VS1": 21, "VS2": 20, "SI1": 17, "SI2": 14, "SI3": 12, "I1": 11, "I2": 11, "I3": 8}, "I": {"IF": 23, "VVS1": 20, "VVS2": 19, "VS1": 18, "VS2": 17, "SI1": 15, "SI2": 13, "SI3": 11, "I1": 10, "I2": 10, "I3": 8}, "J": {"IF": 19, "VVS1": 17, "VVS2": 16, "VS1": 15, "VS2": 14, "SI1": 13, "SI2": 12, "SI3": 11, "I1": 10, "I2": 10, "I3": 7}, "K": {"IF": 16, "VVS1": 15, "VVS2": 14, "VS1": 13, "VS2": 12, "SI1": 11, "SI2": 11, "SI3": 10, "I1": 9, "I2": 9, "I3": 7}, "L": {"IF": 15, "VVS1": 14, "VVS2": 13, "VS1": 12, "VS2": 11, "SI1": 11, "SI2": 10, "SI3": 10, "I1": 9, "I2": 8, "I3": 6}, "M": {"IF": 14, "VVS1": 13, "VVS2": 12, "VS1": 11, "VS2": 10, "SI1": 10, "SI2": 9, "SI3": 9, "I1": 9, "I2": 7, "I3": 5}}},
    {"min": 0.70, "max": 0.89, "data": {"D": {"IF": 70, "VVS1": 56, "VVS2": 44, "VS1": 38, "VS2": 33, "SI1": 29, "SI2": 25, "SI3": 23, "I1": 21, "I2": 19, "I3": 12}, "E": {"IF": 57, "VVS1": 49, "VVS2": 41, "VS1": 36, "VS2": 31, "SI1": 27, "SI2": 23, "SI3": 21, "I1": 19, "I2": 18, "I3": 11}, "F": {"IF": 50, "VVS1": 44, "VVS2": 39, "VS1": 34, "VS2": 29, "SI1": 25, "SI2": 21, "SI3": 19, "I1": 18, "I2": 17, "I3": 11}, "G": {"IF": 42, "VVS1": 37, "VVS2": 34, "VS1": 31, "VS2": 27, "SI1": 23, "SI2": 20, "SI3": 18, "I1": 17, "I2": 16, "I3": 10}, "H": {"IF": 34, "VVS1": 30, "VVS2": 28, "VS1": 26, "VS2": 24, "SI1": 21, "SI2": 18, "SI3": 17, "I1": 16, "I2": 15, "I3": 9}, "I": {"IF": 29, "VVS1": 26, "VVS2": 24, "VS1": 23, "VS2": 21, "SI1": 18, "SI2": 16, "SI3": 15, "I1": 14, "I2": 14, "I3": 9}, "J": {"IF": 24, "VVS1": 22, "VVS2": 20, "VS1": 19, "VS2": 18, "SI1": 17, "SI2": 15, "SI3": 14, "I1": 13, "I2": 13, "I3": 8}, "K": {"IF": 22, "VVS1": 20, "VVS2": 18, "VS1": 17, "VS2": 16, "SI1": 15, "SI2": 14, "SI3": 13, "I1": 12, "I2": 11, "I3": 8}, "L": {"IF": 20, "VVS1": 18, "VVS2": 16, "VS1": 15, "VS2": 14, "SI1": 13, "SI2": 12, "SI3": 12, "I1": 12, "I2": 9, "I3": 7}, "M": {"IF": 18, "VVS1": 16, "VVS2": 14, "VS1": 13, "VS2": 12, "SI1": 12, "SI2": 11, "SI3": 11, "I1": 11, "I2": 8, "I3": 6}}},
    {"min": 0.90, "max": 0.99, "data": {"D": {"IF": 104, "VVS1": 89, "VVS2": 67, "VS1": 57, "VS2": 49, "SI1": 40, "SI2": 32, "SI3": 28, "I1": 27, "I2": 22, "I3": 15}, "E": {"IF": 90, "VVS1": 77, "VVS2": 62, "VS1": 52, "VS2": 45, "SI1": 36, "SI2": 29, "SI3": 26, "I1": 25, "I2": 21, "I3": 14}, "F": {"IF": 82, "VVS1": 71, "VVS2": 57, "VS1": 48, "VS2": 42, "SI1": 34, "SI2": 26, "SI3": 24, "I1": 23, "I2": 20, "I3": 13}, "G": {"IF": 65, "VVS1": 57, "VVS2": 49, "VS1": 43, "VS2": 39, "SI1": 31, "SI2": 25, "SI3": 23, "I1": 22, "I2": 19, "I3": 12}, "H": {"IF": 50, "VVS1": 46, "VVS2": 42, "VS1": 37, "VS2": 34, "SI1": 29, "SI2": 24, "SI3": 23, "I1": 22, "I2": 18, "I3": 12}, "I": {"IF": 45, "VVS1": 41, "VVS2": 37, "VS1": 33, "VS2": 30, "SI1": 26, "SI2": 22, "SI3": 21, "I1": 20, "I2": 17, "I3": 11}, "J": {"IF": 37, "VVS1": 34, "VVS2": 31, "VS1": 28, "VS2": 26, "SI1": 23, "SI2": 21, "SI3": 20, "I1": 19, "I2": 16, "I3": 10}, "K": {"IF": 32, "VVS1": 30, "VVS2": 28, "VS1": 25, "VS2": 23, "SI1": 21, "SI2": 19, "SI3": 18, "I1": 17, "I2": 15, "I3": 9}, "L": {"IF": 27, "VVS1": 25, "VVS2": 23, "VS1": 21, "VS2": 19, "SI1": 18, "SI2": 17, "SI3": 17, "I1": 16, "I2": 13, "I3": 8}, "M": {"IF": 24, "VVS1": 22, "VVS2": 20, "VS1": 19, "VS2": 17, "SI1": 16, "SI2": 15, "SI3": 15, "I1": 14, "I2": 11, "I3": 7}}},
    {"min": 1.00, "max": 1.49, "data": {"D": {"IF": 160, "VVS1": 128, "VVS2": 97, "VS1": 83, "VS2": 69, "SI1": 52, "SI2": 41, "SI3": 36, "I1": 33, "I2": 25, "I3": 16}, "E": {"IF": 125, "VVS1": 111, "VVS2": 88, "VS1": 75, "VS2": 62, "SI1": 48, "SI2": 38, "SI3": 33, "I1": 31, "I2": 24, "I3": 15}, "F": {"IF": 107, "VVS1": 97, "VVS2": 80, "VS1": 68, "VS2": 56, "SI1": 45, "SI2": 35, "SI3": 31, "I1": 29, "I2": 23, "I3": 14}, "G": {"IF": 82, "VVS1": 74, "VVS2": 67, "VS1": 59, "VS2": 51, "SI1": 41, "SI2": 33, "SI3": 29, "I1": 27, "I2": 22, "I3": 13}, "H": {"IF": 61, "VVS1": 56, "VVS2": 52, "VS1": 49, "VS2": 46, "SI1": 38, "SI2": 31, "SI3": 28, "I1": 26, "I2": 21, "I3": 13}, "I": {"IF": 52, "VVS1": 47, "VVS2": 44, "VS1": 41, "VS2": 38, "SI1": 34, "SI2": 29, "SI3": 26, "I1": 24, "I2": 20, "I3": 12}, "J": {"IF": 43, "VVS1": 39, "VVS2": 36, "VS1": 33, "VS2": 31, "SI1": 28, "SI2": 25, "SI3": 23, "I1": 22, "I2": 19, "I3": 12}, "K": {"IF": 36, "VVS1": 33, "VVS2": 31, "VS1": 29, "VS2": 27, "SI1": 25, "SI2": 23, "SI3": 22, "I1": 21, "I2": 18, "I3": 11}, "L": {"IF": 31, "VVS1": 28, "VVS2": 26, "VS1": 25, "VS2": 23, "SI1": 21, "SI2": 20, "SI3": 19, "I1": 18, "I2": 17, "I3": 10}, "M": {"IF": 27, "VVS1": 25, "VVS2": 24, "VS1": 23, "VS2": 21, "SI1": 19, "SI2": 18, "SI3": 17, "I1": 16, "I2": 16, "I3": 10}}},
    {"min": 1.50, "max": 1.99, "data": {"D": {"IF": 210, "VVS1": 187, "VVS2": 154, "VS1": 134, "VS2": 120, "SI1": 93, "SI2": 75, "SI3": 66, "I1": 55, "I2": 35, "I3": 18}, "E": {"IF": 188, "VVS1": 173, "VVS2": 143, "VS1": 122, "VS2": 110, "SI1": 86, "SI2": 68, "SI3": 60, "I1": 52, "I2": 33, "I3": 17}, "F": {"IF": 164, "VVS1": 153, "VVS2": 132, "VS1": 114, "VS2": 103, "SI1": 81, "SI2": 64, "SI3": 57, "I1": 49, "I2": 32, "I3": 16}, "G": {"IF": 136, "VVS1": 126, "VVS2": 114, "VS1": 99, "VS2": 89, "SI1": 75, "SI2": 60, "SI3": 54, "I1": 46, "I2": 30, "I3": 15}, "H": {"IF": 108, "VVS1": 100, "VVS2": 91, "VS1": 81, "VS2": 74, "SI1": 66, "SI2": 55, "SI3": 50, "I1": 42, "I2": 29, "I3": 15}, "I": {"IF": 87, "VVS1": 81, "VVS2": 73, "VS1": 68, "VS2": 63, "SI1": 56, "SI2": 51, "SI3": 46, "I1": 39, "I2": 27, "I3": 14}, "J": {"IF": 74, "VVS1": 67, "VVS2": 61, "VS1": 57, "VS2": 53, "SI1": 48, "SI2": 43, "SI3": 39, "I1": 35, "I2": 26, "I3": 14}, "K": {"IF": 63, "VVS1": 56, "VVS2": 51, "VS1": 47, "VS2": 44, "SI1": 40, "SI2": 37, "SI3": 34, "I1": 31, "I2": 24, "I3": 13}, "L": {"IF": 53, "VVS1": 47, "VVS2": 43, "VS1": 40, "VS2": 38, "SI1": 35, "SI2": 33, "SI3": 31, "I1": 29, "I2": 23, "I3": 12}, "M": {"IF": 46, "VVS1": 41, "VVS2": 39, "VS1": 36, "VS2": 34, "SI1": 32, "SI2": 30, "SI3": 28, "I1": 27, "I2": 22, "I3": 12}}},
    {"min": 2.00, "max": 2.99, "data": {"D": {"IF": 330, "VVS1": 275, "VVS2": 235, "VS1": 205, "VS2": 175, "SI1": 141, "SI2": 113, "SI3": 95, "I1": 80, "I2": 41, "I3": 19}, "E": {"IF": 270, "VVS1": 245, "VVS2": 210, "VS1": 190, "VS2": 160, "SI1": 132, "SI2": 105, "SI3": 88, "I1": 76, "I2": 39, "I3": 18}, "F": {"IF": 245, "VVS1": 220, "VVS2": 195, "VS1": 175, "VS2": 150, "SI1": 123, "SI2": 98, "SI3": 83, "I1": 72, "I2": 37, "I3": 17}, "G": {"IF": 205, "VVS1": 185, "VVS2": 165, "VS1": 150, "VS2": 135, "SI1": 112, "SI2": 92, "SI3": 77, "I1": 68, "I2": 35, "I3": 16}, "H": {"IF": 165, "VVS1": 150, "VVS2": 135, "VS1": 125, "VS2": 115, "SI1": 104, "SI2": 86, "SI3": 71, "I1": 65, "I2": 33, "I3": 15}, "I": {"IF": 135, "VVS1": 120, "VVS2": 110, "VS1": 100, "VS2": 93, "SI1": 86, "SI2": 78, "SI3": 66, "I1": 61, "I2": 31, "I3": 15}, "J": {"IF": 109, "VVS1": 99, "VVS2": 91, "VS1": 84, "VS2": 76, "SI1": 69, "SI2": 63, "SI3": 57, "I1": 54, "I2": 29, "I3": 14}, "K": {"IF": 91, "VVS1": 83, "VVS2": 76, "VS1": 70, "VS2": 63, "SI1": 57, "SI2": 53, "SI3": 50, "I1": 47, "I2": 28, "I3": 14}, "L": {"IF": 78, "VVS1": 71, "VVS2": 66, "VS1": 61, "VS2": 54, "SI1": 50, "SI2": 46, "SI3": 43, "I1": 40, "I2": 27, "I3": 13}, "M": {"IF": 68, "VVS1": 63, "VVS2": 57, "VS1": 54, "VS2": 48, "SI1": 45, "SI2": 42, "SI3": 40, "I1": 38, "I2": 26, "I3": 13}}}
]

def log_to_excel(weight, color, clarity, rate, total):
    file_path = "Diamond_Inventory.xlsx"
    try:
        if not os.path.exists(file_path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Timestamp", "Weight (ct)", "Color", "Clarity", "Rate/ct ($)", "Total ($)"])
            wb.save(file_path)
        
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        ws.append([datetime.now().strftime("%Y-%m-%d %H:%M"), weight, color, clarity, rate, total])
        wb.save(file_path)
        return True
    except PermissionError:
        return "LOCKED"
    except Exception as e:
        return str(e)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/calculate', methods=['POST'])
def calculate():
    try:
        data = request.get_json()
        weight = float(data.get('weight', 0))
        color = data.get('color', '').upper()
        clarity = data.get('clarity', '').upper()

        # Find the correct table based on weight
        table = next((t for t in RAP_DATA if t["min"] <= weight <= t["max"]), None)

        if table and color in table["data"] and clarity in table["data"][color]:
            rate_per_ct = table["data"][color][clarity] * 100
            total_value = round(rate_per_ct * weight, 2)
            
            save_status = log_to_excel(weight, color, clarity, rate_per_ct, total_value)
            
            return jsonify({
                "success": True,
                "rate": rate_per_ct,
                "total": total_value,
                "saved": save_status
            })
        else:
            return jsonify({"success": False, "error": "Weight/Grade out of Rapaport range."})
            
    except Exception as e:
        return jsonify({"success": False, "error": f"Invalid Data: {str(e)}"})

@app.route('/open-file')
def open_file():
    path = "Diamond_Inventory.xlsx"
    if os.path.exists(path):
        if platform.system() == "Windows": os.startfile(path)
        else: subprocess.call(["open", path])
        return jsonify({"success": True})
    return jsonify({"success": False, "error": "No file created yet."})

if __name__ == '__main__':
    app.run(debug=True, port=5000)