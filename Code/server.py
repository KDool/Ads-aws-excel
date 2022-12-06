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

@app.route('/setup-manual',methods=['GET', 'POST'])
def setup_manual():
    # return redirect('/user')
    cam_id = ''
    asin_list = []
    budget = ''
    productName = ''
    bid=''
    billing_strategy = ''
    if request.method == 'POST':
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
        output_df = setup.createDataFrame_asin(cam_id=cam_id,asin_list=asin_list,budget=budget,productName=product_name,bid=bid,billing_strategy=billing_strategy)
        #     # Creating output and writer (pandas excel writer)
        out = BytesIO()
        writer = pd.ExcelWriter(out, engine='xlsxwriter')

        # Export data frame to excel
        output_df.to_excel(excel_writer=writer, index=False, sheet_name='Sheet1')
        writer.save()
        writer.close()
        out.seek(0)
        # return render_template('./home.html')
        return send_file(out, attachment_filename="testing.xlsx", as_attachment=True)
    else:
        return render_template('./home.html')
    # return send_file(out, attachment_filename="testing.xlsx", as_attachment=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080,debug=True)
