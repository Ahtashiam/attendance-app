from flask import Flask, render_template, request, send_file
from datetime import datetime
from openpyxl import Workbook
import os

app = Flask(__name__)
entries = []  # Temporary storage

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        entry = [
            request.form['name'],
            request.form['designation'],
            request.form['attendance'],
            datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ]
        entries.append(entry)
    return render_template('index.html')

@app.route('/generate_excel')
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    
    # Write header
    ws.append(['Name', 'Designation', 'Attendance', 'DateTime'])
    
    # Write data rows
    for entry in entries:
        ws.append(entry)

    # Save file
    file_path = "attendance.xlsx"
    wb.save(file_path)
    
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
