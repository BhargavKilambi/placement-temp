from flask import Flask,render_template,request,redirect
from werkzeug import secure_filename
import xlrd
import os
import sys
import json

app = Flask(__name__)

com=['Infosys','Cognizont','Capgemini']
infosys = []
preferences = []
placed = []
@app.route('/home',methods=['GET','POST'])
def index():
    if request.method == 'POST':
        global preferences
        reForm = request.form
        count = reForm['count']
        f1 = request.files['f1']
        f1.save((f1.filename))
        workbook = xlrd.open_workbook(os.path.basename(f1.filename))
        worksheet = workbook.sheet_by_name(workbook.sheet_names()[0])
        data = []
        keys = [v.value for v in worksheet.row(0)]
        for row_number in range(worksheet.nrows):
            if row_number == 0:
                continue
            row_data = {}
            for col_number, cell in enumerate(worksheet.row(row_number)):
                row_data[keys[col_number]] = cell.value
            data.append(row_data)
        preferences = data
        return redirect('/test')
    return render_template('index.html')

@app.route('/test',methods=['GET','POST'])
def test():
    if request.method == 'POST':
        reForm = request.form
        f1 = request.files['f1']
        f1.save((f1.filename))
        workbook = xlrd.open_workbook(os.path.basename(f1.filename))
        worksheet = workbook.sheet_by_name(workbook.sheet_names()[0])
        data = []
        keys = [v.value for v in worksheet.row(0)]
        for row_number in range(worksheet.nrows):
            if row_number == 0:
                continue
            row_data = {}
            for col_number, cell in enumerate(worksheet.row(row_number)):
                row_data[keys[col_number]] = cell.value
                data.append(row_data)
        infosys = data
        for row in preferences:
            result = check_placed(row['Roll No.'],"Infosys",infosys)
            placed.append({'Roll No.':str(row['Roll No.']).split('.')[0],'Result':result})
        with open('out.json', 'w') as json_file:
            json_file.write(json.dumps({'data': placed}))
        return render_template('result.html',placed=placed)
    return render_template('test.html')


def check_placed(number,company_name,students):
    result = "Not Placed"
    for count in (students for student in students if student['Roll No.']==number):
        result = company_name
    return result
    # b=conn2.execute("select count(student_id) from {} where student_id={}".format(company_name,number))
    # count=b.fetchone()
    # if(count[0]>0):
    #     return company_name
    # else:
    #     return "Not Placed"   

if __name__ == '__main__':
    app.run(debug=True)
