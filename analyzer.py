import openpyxl, json
from flask import Flask, request, render_template
from statistics import mean, stdev

app = Flask(__name__)

def get_data(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = {}

    # Calculate column statistics
    for col in sheet.columns:
        values = [cell.value for cell in col if isinstance(cell.value, (int, float))]
        if values:
            col_letter = col[0].column_letter
            data[f"Col {col_letter}"] = {
                "Sum": sum(values),
                "Average": mean(values),
                "Max": max(values),
                "Min": min(values),
                "StdDev": stdev(values)
            }

    # Calculate row statistics
    for row in sheet.rows:
        values = [cell.value for cell in row if isinstance(cell.value, (int, float))]
        if values:
            row_num = row[0].row
            data[f"Row {row_num}"] = {
                "Sum": sum(values),
                "Average": mean(values),
                "Max": max(values),
                "Min": min(values),
                "StdDev": stdev(values)
            }

    return render_template('index.html', data=data)

@app.route('/', methods=['POST'])
def upload_file():
    file = request.files['file']
    return get_data(file)

@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)