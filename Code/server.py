from cgi import test
from logging.handlers import RotatingFileHandler
from unittest import result
from flask import Flask, render_template, request, redirect,send_file,flash,url_for
from graphviz import render
from matplotlib.pyplot import bar_label
import yaml
import sys
import pandas as pd
import setup_manual
import setup_Auto
import setup_SD
import optimize
from io import BytesIO
from werkzeug.utils import secure_filename
import os




UPLOAD_FOLDER = '../temp_uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

# app = Flask(__name__)
app = Flask(__name__,template_folder='template')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER



  # List of SetUp dataframe manuals -- Recreate when Total submit button action
global setup_dataframes_list
setup_dataframes_list = []

@app.route('/setup-manual',methods=['GET', 'POST'])
# global setup_dataframes_list
def setup_bulk():
    # return redirect('/user')
    global setup_dataframes_list
    # setup_dataframes_list = []

    if request.method == 'POST' and request.form['setup'] == 'sp_manual':
        f = request.files['file_input']
        input_df = pd.read_excel(f,index_col=False)
        print('SP Manual')
        print(input_df)
        bulk_dataframe = setup_manual.create_bulk_dataframe(input_df)
        print(bulk_dataframe)
        out = BytesIO()
        writer = pd.ExcelWriter(out, engine='xlsxwriter')
        # Export data frame to excel
        bulk_dataframe.to_excel(excel_writer=writer, index=False, sheet_name='Sheet1')
        writer.save()
        writer.close()
        out.seek(0)
        return send_file(out, attachment_filename="testing.xlsx", as_attachment=True)
        # return send_file(out, download_name="testing.xlsx", as_attachment=True)
    elif request.method == 'POST' and request.form['setup'] == 'sp_auto':
        f = request.files['file_input']
        input_df = pd.read_excel(f,index_col=False)
        print('SP Auto')
        print(input_df)
        bulk_dataframe = setup_Auto.createResultDataFrame(input_df)
        print(bulk_dataframe)
        out = BytesIO()
        writer = pd.ExcelWriter(out, engine='xlsxwriter')
        # Export data frame to excel
        bulk_dataframe.to_excel(excel_writer=writer, index=False, sheet_name='Sheet1')
        writer.save()
        writer.close()
        out.seek(0)
        return send_file(out, attachment_filename="testing.xlsx", as_attachment=True)
    elif request.method == 'POST' and request.form['setup'] == 'sd':
        f = request.files['file_input']
        input_df = pd.read_excel(f,index_col=False)
        print('SD')
        print(input_df)
        bulk_dataframe = setup_SD.createResultDataFrame(input_df)
        print(bulk_dataframe)
        out = BytesIO()
        writer = pd.ExcelWriter(out, engine='xlsxwriter')
        # Export data frame to excel
        bulk_dataframe.to_excel(excel_writer=writer, index=False, sheet_name='Sheet1')
        writer.save()
        writer.close()
        out.seek(0)
        return send_file(out, attachment_filename="testing.xlsx", as_attachment=True)
    else:
        return render_template('./setup_manual_asin.html')
    # return send_file(out, attachment_filename="testing.xlsx", as_attachment=True)


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/upload')
def upload_file():
   return render_template('upload.html')
	
@app.route('/uploader', methods = ['GET', 'POST'])
def uploader_file():
    if request.method == 'POST':
        file_bulk = request.files['file_bulk']
        file_report = request.files['file_report']
        # f.save(secure_filename(f.filename))

        acos = float(request.form['Acos'])
        clicks = int(request.form['Clicks'])
        spend = float(request.form['Spend'])
        print(acos,clicks,spend)
        df_CSP = optimize.read_CSP_report(excel_file=file_report)
        if request.form['sp_type']=='sp_products':
            df_CSP_filter = optimize.filter_CSP_negative(CSP_df=df_CSP,acos=acos,clicks=clicks,spend=spend)
            df_bulk = optimize.read_bulk_report_CSP(excel_file=file_bulk)
        elif request.form['sp_type']=='sp_brands':
            df_CSP_filter = optimize.filter_Brands_negative(CSP_df=df_CSP,acos=acos,clicks=clicks,spend=spend)
            df_bulk = optimize.read_bulk_report_Brands(excel_file=file_bulk)
        else:
            return "ERROR"
        # df_bulk = optimize.read_bulk_report(excel_file=file_bulk)
        df_campaign_info = df_bulk[['Campaign Id','Campaign Name']]
        df_campaign_info.drop_duplicates(inplace=True)
        df_campaign_info.reset_index(inplace=True)

        df_cd = optimize.get_campid_toDF(filtered_df=df_CSP_filter,bulk_df=df_campaign_info)
        df_export = optimize.export_excel_files(filtered_df=df_cd)
        
        out = BytesIO()
        writer = pd.ExcelWriter(out, engine='xlsxwriter')
        # Export data frame to excel
        df_export.to_excel(excel_writer=writer, index=False, sheet_name='Sheet1')
        writer.save()
        writer.close()
        out.seek(0)

        return send_file(out, attachment_filename="testing.xlsx", as_attachment=True)
    # return 'Uploaded sucessfully'






if __name__ == '__main__':
    # setup_dataframes_list = []
    app.run(host='0.0.0.0', port=8080,debug=True)
