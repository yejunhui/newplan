import pandas as pd
from flask import Flask,render_template
app = Flask(__name__)

@app.route('/',methods=['GET','POST'])
def login():
    d={}
    d['title'] = 'Login'
    d['data'] = [[1,2,3],[2,3,4],[3,4,5],[4,5,6]]
    return render_template('login.html',d=d)

if __name__ == '__main__':
    app.run(debug=True,host='0.0.0.0',port=80)