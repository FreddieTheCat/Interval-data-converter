from flask import Flask,render_template,request,send_file
from flask_sqlalchemy import SQLAlchemy
import os
import pandas as pd
from openpyxl import load_workbook
import sqlalchemy as db

######### function#######################
def transform(df):
     #count the number of columns in the data frame
    col=len(df.columns)
    if col>3:
    #Transform a matrix to a vector
        df=df.set_index(df.columns[0]).stack().reset_index()
        df[['Date','level_1']]=df[['Date','level_1']].astype(str)
        df['dtime']=df['Date']+' '+df['level_1']
        df['dtime'] = pd.to_datetime(df['dtime'])
        df=df.drop(['Date', 'level_1'], axis=1)
        df.columns=['KW','dtime']
        df=df[['dtime','KW']]
        df=df.sort_values(by='dtime',ascending=False)
        df.reset_index(inplace=True,drop=True)
    #df.index = pd.to_datetime(df[df.columns[[0,1]]].astype(str).apply('-'.join,1))
    else:
        df.columns =(['dtime','kW'])
        df['dtime'] = pd.to_datetime(df['dtime'])
        df['dtime'] = pd.to_datetime(df['dtime'])
#find the interval by substracting the second date from the first one
    a = df.loc[0, 'dtime']
    b = df.loc[1, 'dtime']
    c = a - b
    minutes = c.total_seconds() / 60
    d=int(minutes) #d can be only 15 ,30 or 60
    #This function will create new row to the time series anytime when it finds gaps and will fill it with NaN or leave it blank.
    #df.drop_duplicates(keep='first') keeps the first value of duplicates
    if d==15:
        df.drop_duplicates(keep='first',inplace=True)
        df= df.set_index('dtime').asfreq('-15T')
    elif d==30:
        df.drop_duplicates(keep='first',inplace=True)
        df= df.set_index('dtime').asfreq('-30T')
    elif d==60:
        df.drop_duplicates(keep='first',inplace=True)
        df= df.set_index('dtime').asfreq('-60T')
    else:
        None
    return df

###########Flask APP######################
app=Flask(__name__)

#app.config['Save_File']='C:\Users\......'
#db=SQLAlchemy(app)

#class FileContents(db.Model):
#    id=db.Column(db.Integer,primary_key=True)
#    name=db.Column(db.String(300))
#    data=db.Column(db.LargeBinary)


@app.route('/')
def index():
    return render_template('firstpage.html')


@app.route('/upload',methods=['Get','POST'])
def upload():
    file=request.files['inputfile']

    xls=pd.ExcelFile(file)

    name_dict = {}

    snames = xls.sheet_names

    for sn in snames:
        name_dict[sn] = xls.parse(sn)

    for key, value in name_dict.items():
        transform(value)
    transformed_dict={}

    for key, value in name_dict.items():
        transformed_dict[key]=transform(value)

    #### wirte to excel example:
    writer = pd.ExcelWriter("MyData.xlsx", engine='xlsxwriter')
    for name, df in transformed_dict.items():
        df.to_excel(writer, sheet_name=name)
    writer.save()



if __name__=='__main__':
    app.run(port=5000)
