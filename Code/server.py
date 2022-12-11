from cgi import test
from logging.handlers import RotatingFileHandler
from unittest import result
from flask import Flask, render_template, request, redirect,send_file,flash,url_for
from graphviz import render
from matplotlib.pyplot import bar_label
import yaml
import sys
import pandas as pd
import setup
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
def setup_manual():
    # return redirect('/user')
    global setup_dataframes_list
    # setup_dataframes_list = []
    cam_id = ''
    asin_list = []
    budget = ''
    product_name = ''
    bid=''
    billing_strategy = ''
    if request.method == 'POST' and request.form['button_action'] == 'Asin Append':
        cam_id = request.form['camp_name']
        asin_list = request.form['asin_list'].split(',')
        budget = request.form['budget']
        product_name = request.form['product_name']
        bid = request.form['bid']
        billing_strategy = request.form['billing_strategy']

        print('cam id: ',cam_id)
        print('Asin_list: ',asin_list)
        print('budget: ',budget)
        print('Product Name: ',product_name)
        print('Bid: ',bid)
        print('Billing Strategy: ',billing_strategy)
        asin_df = setup.createDataFrame_asin(cam_id=cam_id,asin_list=asin_list,budget=budget,productName=product_name,bid=bid,billing_strategy=billing_strategy)
        print('ASIN DF ==============================================: ',asin_df)
        setup_dataframes_list.append(asin_df)
        print('SET UP DATA LIST============================: ', setup_dataframes_list)

        return render_template('./setup_manual_asin.html')
        # return send_file(out, download_name="testing.xlsx", as_attachment=True)
    elif request.method == 'POST' and request.form['button_action'] == 'Keyword Append':
        cam_id = request.form['camp_name']
        keyword_list = request.form['keyword_list'].split(',')
        budget = request.form['budget']
        product_name = request.form['product_name']
        bid = request.form['bid']
        billing_strategy = request.form['billing_strategy']
        match_type = request.form['match_type']
        portfolio_id = request.form['portfolio_id']

        print('cam id: ',cam_id)
        print('keyword_list: ',keyword_list)
        print('budget: ',budget)
        print('Product Name: ',product_name)
        print('Bid: ',bid)
        print('Billing Strategy: ',billing_strategy)
        print('Portfolio_Id: ',portfolio_id)
        print('Match Type: ',match_type)
        
        kw_df = setup.createDataFrame_keyword(cam_id=cam_id,keyword_list=keyword_list,
                                              budget=budget,productName=product_name,bid=bid,
                                              billing_strategy=billing_strategy,portfolio_id=portfolio_id,
                                              match_type=match_type)
        # print('Keyword DF: ',kw_df)
        
        setup_dataframes_list.append(kw_df)
        
        # print("")
        print('SET UP DATA LIST============================: ', setup_dataframes_list)
        return render_template('./setup_manual_asin.html')
    elif request.method == 'POST' and request.form['button_action'] == 'Total Submit':
        print('SET UP DATA LIST============================: ', setup_dataframes_list)
        total_df = pd.concat(setup_dataframes_list)
        out = BytesIO()
        writer = pd.ExcelWriter(out, engine='xlsxwriter')
        # Export data frame to excel
        total_df.to_excel(excel_writer=writer, index=False, sheet_name='Sheet1')
        writer.save()
        writer.close()
        out.seek(0)
        
        setup_dataframes_list = [] #### Clean DataFrame
        # return send_file(out, download_name="testing.xlsx", as_attachment=True)
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
        f = request.files['file']
        f.save(secure_filename(f.filename))

        acos = float(request.form['Acos'])
        clicks = int(request.form['Clicks'])
        spend = float(request.form['Spend'])
        print(acos,clicks,spend)

        df_CSP = optimize.read_CSP_report(excel_file=f,sheet_name='Sponsored Product Search Term R')
        df_CSP_filter = optimize.filter_CSP_negative(CSP_df=df_CSP,acos=acos,clicks=clicks,spend=spend)
        df_cd = optimize.get_campid_toDF(filtered_df=df_CSP_filter,table_name='table_bulk_products')
        df_export = optimize.export_excel_files(filtered_df=df_cd)
        
        out = BytesIO()
        writer = pd.ExcelWriter(out, engine='xlsxwriter')
        # Export data frame to excel
        df_export.to_excel(excel_writer=writer, index=False, sheet_name='Sheet1')
        writer.save()
        writer.close()
        out.seek(0)

        return send_file(out, attachment_filename="testing.xlsx", as_attachment=True)
    #   print(df)
    # return 'Uploaded sucessfully'

# @app


if __name__ == '__main__':
    # setup_dataframes_list = []
    app.run(host='0.0.0.0', port=8080,debug=True)
