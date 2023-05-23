from flask import Flask,jsonify,json,redirect,url_for
from flask import request , send_file , after_this_request
from flask import render_template
import os
import xlsxwriter

app = Flask(__name__)
app.secret_key = os.urandom(24)



@app.route('/')
def home():
   return "Successfully Displayed Python App Service"


@app.route('/excel')
def excel():
    date = "23-05-2023"
    workbook = xlsxwriter.Workbook(f'Attendance_{date}.xlsx')
    worksheet = workbook.add_worksheet(date)
    #Excel Formatting
    bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    text_wrap = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    center = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    
    #Setting Column Width
    worksheet.set_column("C:C",10)
    worksheet.set_column("D:D",5)
    worksheet.set_column("E:E",50)
    worksheet.set_column("F:F",15)
    worksheet.set_column("G:G",15)
    worksheet.set_column("H:H",15)
    worksheet.set_column("I:I",20)
    
    #Adding Data to 3rd Row
    worksheet.write('C3', date , bold)
    worksheet.write('D3', 'Total', bold)
    worksheet.write('E3', 'At Office', bold)
    worksheet.write('F3', 'Work From Home', bold)
    worksheet.write('G3', 'At Customer Site', bold)
    worksheet.write('H3', 'On Leave Not Sick', bold)
    worksheet.write('I3', 'Sick', bold)
    
    #Adding Data to 4th Row
    worksheet.write('E4',"27",center)
    worksheet.write('H4',"3",center)
    worksheet.write('I4',"2",center)
    
    data = ['Ranimounika', 'Ranimounika', 'Ranimounika']
    
    #Adding Data to 5th Row
    worksheet.write('D5',"82",center)
    worksheet.write('E5',"Amit P, Arunagiri S, Srinivas Jonnala, Balachandra C, Rajalaxmi Pillai, Ganesh A, Sreevidya, Swethanjali, Savita Walikar,Aparijitha, Laxmi AS, Sushitha, Jyothi S,  Natesha, Pooja R, Rithik, Sambhu ,Raghavendra, Syamkumar, Tejaswini, Chethana, Triveni,  Ashwini Divatar, Aruna, Ajay M, Arun K, Deepa B",text_wrap)    
    worksheet.write("F5","50",center)
    worksheet.write("H5",','.join(data),text_wrap)
    worksheet.write("I5","Ravipriya, Sivanunna",text_wrap)

    
    
    
    workbook.close()
    return "excel man"



if __name__ == '__main__':
   app.run(host='0.0.0.0',port=80)