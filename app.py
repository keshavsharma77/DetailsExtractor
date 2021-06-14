import os
import docx2txt
import PyPDF2

from docx2pdf import convert
from flask import Flask, flash, request, redirect, render_template, url_for, Response
from werkzeug.utils import secure_filename
#from flask_sqlalchemy import SQLAlchemy
from flask_mysqldb import MySQL
#from flaskext.mysql import MySQL
import io
import xlwt
#import pymysql
import yaml
import re

#db = pymysql.connect("localhost", "root", "1234", "pdfdetails")
'''db = pymysql.connect(host='localhost',
                             user='root',
                             password='1234',
                             database='pdfdetails',
                             port = 3307,
                             charset='utf8mb4',
                             cursorclass=pymysql.cursors.DictCursor)'''


app = Flask(__name__)
'''app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root:@localhost/pdfdetail'
db = SQLAlchemy(app)
SQLALCHEMY_TRACK_MODIFICATIONS = False

class details(db.Model):
    sno = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(40), nullable=False)
    contact = db.Column(db.String(13), nullable=False)
    location = db.Column(db.String(100), nullable=False)
'''

# Configure db
db = yaml.load(open('db.yaml'))
app.config['MYSQL_HOST'] = db['mysql_host']
app.config['MYSQL_USER'] = db['mysql_user']
app.config['MYSQL_PASSWORD'] = db['mysql_password']
app.config['MYSQL_DB'] = db['mysql_db']
app.config['MYSQL_PORT'] = db['mysql_port']

mysql = MySQL(app)

app.secret_key = "secret key"
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Get current path
path = os.getcwd()
# file Upload
UPLOAD_FOLDER = os.path.join(path, 'uploads')

# Make directory if uploads is not exists
if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Allowed extension you can set your own
ALLOWED_EXTENSIONS = set(['doc', 'pdf', 'docx'])
CONVERT_EXTENSIONS = set(['doc', 'docx'])

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in CONVERT_EXTENSIONS



@app.route("/")
def upload_form():
    return render_template('index.html')


@app.route('/', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        if 'files[]' not in request.files:
            flash('No file part')
            return redirect(request.url)

        files = request.files.getlist('files[]')

        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                if file and convert_file(file.filename):
                    file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                    #convert('uploads/filename')
                    convert("uploads")
                    my_text = docx2txt.process(file)

                    pattern = re.compile(r'[0-9a-zA-Z]+@[0-9a-zA-Z]+\.[0-9a-zA-Z]+')
                    pattern1 = re.compile(r'[0-9]{10}')
                    matches = pattern.finditer(my_text)
                    matches1 = pattern1.finditer(my_text)
                    for match in matches:
                        email1 = match.group(0)
                        print(email1)
                    for match1 in matches1:
                        contact1 = match1.group(0)
                        print(contact1)
                    # = url_for('uploaded_file', filename=filename)
                    lcn = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    print(lcn)
                    cur = mysql.connection.cursor()
                    # cursor = db.cursor()
                    cur.execute("INSERT INTO details(email, contact, location) VALUES(%s, %s, %s)",
                                (email1, contact1, lcn))
                    # sql = "INSERT INTO details(email, contact, location) VALUES(%s, %s, %s)"
                    # cursor.execute(sql, (email1, contact1, lcn))
                    mysql.connection.commit()
                    cur.close()

                else:
                    file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                    lcn1 = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    pdffileobj = open(str(lcn1), 'rb')
                    pdfreader = PyPDF2.PdfFileReader(pdffileobj)
                    pageobj = pdfreader.getPage(0)
                    pjh = ""
                    pjh = pageobj.extractText()
                    pattern = re.compile(r'[0-9a-zA-Z]+@[0-9a-zA-Z]+\.[0-9a-zA-Z]+')
                    pattern1 = re.compile(r'[0-9]{10}')
                    matches = pattern.finditer(pjh)
                    matches1 = pattern1.finditer(pjh)
                    for match in matches:
                        email2 = match.group(0)
                        print(email2)
                    for match1 in matches1:
                        contact2 = match1.group(0)
                        print(contact2)
                    # = url_for('uploaded_file', filename=filename)
                    print(lcn1)
                    pdffileobj.close()
                    cur = mysql.connection.cursor()
                    # cursor = db.cursor()
                    cur.execute("INSERT INTO details(email, contact, location) VALUES(%s, %s, %s)",
                                (email2, contact2, lcn1))
                    # sql = "INSERT INTO details(email, contact, location) VALUES(%s, %s, %s)"
                    # cursor.execute(sql, (email1, contact1, lcn1))
                    mysql.connection.commit()
                    cur.close()



                    #db.commit()
                '''
                entry = details(email=str(email1), contact=str(contact1), location=str(lcn))
                db.session.add(entry)
                db.session.commit()'''

        flash('File(s) successfully uploaded')
        return redirect('/about')

#from Model import details


@app.route("/about")
def harry():
    '''
    cursor = db.cursor()
    sql = "SELECT * FROM details"
    cursor.execute(sql)
    results = cursor.fetchall()
    '''
    cur = mysql.connection.cursor()
    resultValue = cur.execute("SELECT * FROM details")
    if resultValue > 0:
        results = cur.fetchall()
    return render_template('about.html', results=results)


@app.route('/download')
def download():
    return render_template('download.html')

@app.route('/download/report/excel')
def download_report():
    cur = mysql.connection.cursor()

    cur.execute("SELECT email, contact, location FROM details")
    result = cur.fetchall()

    # output in bytes
    output = io.BytesIO()
    # create WorkBook object
    workbook = xlwt.Workbook()
    # add a sheet
    sh = workbook.add_sheet('Information Report')

    # add headers
    sh.write(0, 0, 'Email')
    sh.write(0, 1, 'Contact')
    sh.write(0, 2, 'Location')

    idx = 0
    for row in result:
        sh.write(idx + 1, 0, row[0])
        sh.write(idx + 1, 1, row[1])
        sh.write(idx + 1, 2, row[2])
        idx += 1

    workbook.save(output)
    output.seek(0)

    return Response(output, mimetype="application/ms-excel",
                    headers={"Content-Disposition": "attachment;filename=employee_report.xls"})


app.run(debug=True)

if __name__ == "__main__":
    app.run(host='127.0.0.1',port=5000,debug=False,threaded=True)

