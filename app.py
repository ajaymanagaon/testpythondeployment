from flask import Flask,jsonify,json,redirect,url_for
from flask import request , send_file , after_this_request
from flask import render_template
import os

app = Flask(__name__)
app.secret_key = os.urandom(24)



@app.route('/')
def home():
   return "Successfully Displayed Python App Service"



if __name__ == '__main__':
   app.run(host='0.0.0.0',port=80)