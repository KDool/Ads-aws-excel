from cgi import test
from logging.handlers import RotatingFileHandler
from unittest import result
from flask import Flask, render_template, request, redirect,send_file
from graphviz import render
from matplotlib.pyplot import bar_label
import yaml
import sys
import pandas as pd
import setup
from io import BytesIO

app = Flask(__name__,template_folder='template')

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


if __name__ == '__main__':
    # setup_dataframes_list = []
    app.run(host='0.0.0.0', port=8080,debug=True)
