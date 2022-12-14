import pandas as pd
import openpyxl
import xlsxwriter
from datetime import date

def create_single_row_SP(columns:list,cam_id):
    single_row_dictionary = {}
    for item in columns:
        if item == 'Product':
            single_row_dictionary[item] = 'Sponsored Products'
        elif item == 'Operation':
            single_row_dictionary[item] = 'Create'
        elif item == 'Campaign Id':
            single_row_dictionary[item] = cam_id
        elif item == 'State':
            single_row_dictionary[item] = 'enabled'
        else:
            single_row_dictionary[item] = None
    return single_row_dictionary

def createDataFrameSP_asin(cam_id,asin_list:list,budget,productName,bid_list:list,billing_strategy,placement,percentage):
    columns=['Product','Entity','Operation','Campaign Id','Ad Group Id','Portfolio Id',
                           'Ad Id','Keyword Id','Product Targeting Id','Campaign Name','Ad Group Name',
                           'Start Date','End Date','Targeting Type','State','Daily Budget','sku','asin',
                           'Ad Group Default Bid','Bid','Keyword Text','Match Type','Bidding Strategy',
                           'Placement','Percentage','Product Targeting Expression']
    df = pd.DataFrame(columns=columns)

# 5 dong co dinh
    for i in range (0,5):
        df = df.append(create_single_row_SP(columns=columns,cam_id=cam_id),ignore_index=True)
        
    df.iloc[0]['Entity'] = 'Campaign'
    df.iloc[0]['Campagin Name'] = cam_id
    df.iloc[0]['Targeting Type'] = 'Manual'
    df.iloc[0]['Start Date'] = date.today().strftime("%Y%m%d")
    df.iloc[0]['Daily Budget'] = budget
    df.iloc[0]['Bidding Strategy'] = billing_strategy

    df.iloc[1]['Entity'] = 'Bidding Adjustment'
    df.iloc[1]['Placement'] = placement
    df.iloc[1]['Percentage'] = percentage

    df.iloc[2]['Entity'] = 'Bidding Adjustment'
    df.iloc[2]['Placement'] = placement
    df.iloc[2]['Percentage'] = percentage

    df.iloc[3]['Entity'] = 'Ad group'
    df.iloc[3]['Ad Group Id'] = cam_id
    df.iloc[3]['Ad Group Name'] = cam_id
    # df.iloc[3]['Ad Group Default Bid'] = bid
    
    df.iloc[4]['Entity'] = 'Product ad'
    df.iloc[4]['Ad Group Id'] = cam_id
    df.iloc[4]['sku'] = productName


# Optional Asin - so dong Asin
    for i in range(len(asin_list)):
        asin_row = create_single_row_SP(columns=columns,cam_id=cam_id)
        asin_row['Entity'] = 'Product targeting'
        asin_row['Ad Group Id'] = cam_id
        if len(bid_list)>1:
            if i < len(bid_list):
                asin_row['Bid'] = bid_list[i]
            else:
                asin_row['Bid'] = bid_list[-1]
        else:
            asin_row['Bid'] = bid_list[0]
        asin_row['Product Targeting Expression'] = '''asin="''' +str(asin_list[i])+ '''"'''
        df = df.append(asin_row,ignore_index=True)

    return df


def createDataFrameSP_keyword(cam_id,keyword_list:list,budget,productName,bid,billing_strategy,portfolio_id,match_type,placement,percentage):
    columns=['Product','Entity','Operation','Campaign Id','Ad Group Id','Portfolio Id',
                           'Ad Id','Keyword Id','Product Targeting Id','Campaign Name','Ad Group Name',
                           'Start Date','End Date','Targeting Type','State','Daily Budget','sku','asin',
                           'Ad Group Default Bid','Bid','Keyword Text','Match Type','Bidding Strategy',
                           'Placement','Percentage','Product Targeting Expression']
    df = pd.DataFrame(columns=columns)

# 5 dong co dinh
    for i in range (0,5):
        df = df.append(create_single_row_SP(columns=columns,cam_id=cam_id),ignore_index=True)
        
    df.iloc[0]['Entity'] = 'Campaign'
    df.iloc[0]['Campagin Name'] = cam_id
    df.iloc[0]['Targeting Type'] = 'Manual'
    df.iloc[0]['Start Date'] = date.today().strftime("%Y%m%d")
    df.iloc[0]['Daily Budget'] = budget
    df.iloc[0]['Bidding Strategy'] = billing_strategy
    df.iloc[0]['Portfolio Id'] = portfolio_id

    df.iloc[1]['Entity'] = 'Bidding Adjustment'
    df.iloc[1]['Placement'] = placement
    df.iloc[1]['Percentage'] = percentage

    df.iloc[2]['Entity'] = 'Bidding Adjustment'
    df.iloc[2]['Placement'] = placement
    df.iloc[2]['Percentage'] = percentage

    df.iloc[3]['Entity'] = 'Ad group'
    df.iloc[3]['Ad Group Id'] = cam_id
    df.iloc[3]['Ad Group Name'] = cam_id
    df.iloc[3]['Ad Group Default Bid'] = bid
    
    df.iloc[4]['Entity'] = 'Product ad'
    df.iloc[4]['Ad Group Id'] = cam_id
    df.iloc[4]['sku'] = productName
    df.iloc[4]['Bid'] = bid

# Optional Product Targetting- so dong Keyword
    for item in keyword_list:
        keyword_row = create_single_row_SP(columns=columns,cam_id=cam_id)
        keyword_row['Entity'] = 'Product targeting'
        keyword_row['Ad Group Id'] = cam_id
        keyword_row['Keyword Text'] = str(item)
        keyword_row['Match Type'] = match_type
        df = df.append(keyword_row,ignore_index=True)
    
    return df

def create_bulk_dataframe(input_df:pd.DataFrame):
    dataframes_list =[] 
    for i in range(len(input_df)):
        cam_id = str(input_df.iloc[i]['CODE']) + ' ' +  str(input_df.iloc[i]['Market']) + ' ' + str(input_df.iloc[i]['PPC Type']) + ' ' + str(input_df.iloc[i]['Match type'])+ ' ' + str(input_df.iloc[i]['BRAND']) +' ' + str(input_df.iloc[i]['Date']) +' ' + str(input_df.iloc[i]['PIC'] + ' ' + str(input_df.iloc[i]['STT']))
        kw_as_list = str(input_df.iloc[i]['Targeting']).split(',')
        if input_df.iloc[i]['Targeting type'] == 'KW':
            kw_df = createDataFrameSP_keyword(cam_id=cam_id,keyword_list=kw_as_list,budget=input_df.iloc[i]['Budget'],productName=input_df.iloc[i]['SKU'],
                                                bid=str(input_df.iloc[i]['Bid']),billing_strategy=input_df.iloc[i]['Bid strategy'],
                                                portfolio_id=input_df.iloc[i]['Portfolio'],match_type=input_df.iloc[i]['Match type'],
                                                placement=str(input_df.iloc[i]['Placement']),percentage="{:.0%}".format(input_df.iloc[i]['Percentage']))
            dataframes_list.append(kw_df)
        elif input_df.iloc[i]['Targeting type'] == 'ASIN':
            bid_list = str(input_df.iloc[i]['Bid']).split(',')
            asin_df = createDataFrameSP_asin(cam_id=cam_id,asin_list=kw_as_list,budget=input_df.iloc[i]['Budget'],productName=input_df.iloc[i]['SKU'],
                                                bid_list=bid_list,billing_strategy=input_df.iloc[i]['Bid strategy'],placement=str(input_df.iloc[i]['Placement']),percentage="{:.0%}".format(input_df.iloc[i]['Percentage']))
            dataframes_list.append(asin_df)
        
    result_df = pd.concat(dataframes_list)
    return result_df

#### WRITE TO EXCEL #######
# df.to_excel('myexcel.xlsx',sheet_name='Sheet1',engine='xlsxwriter')
def export_excel(output_df:pd.DataFrame,file_export_path='',sheet_name=''):
    writer = pd.ExcelWriter(file_export_path, engine='xlsxwriter')
    output_df.to_excel(writer,sheet_name = sheet_name, index=False)
    # output_df.to_excel(writer,sheet_name = 'sheet_keyword', index=False)
    writer.save() 


if __name__ == '__main__':
    output_asin_df = createDataFrameSP_asin(cam_id='ASD123456',asin_list=['Khai1','Khai2'],budget=2.9,productName='Newspaper',bid=0.6,billing_strategy='Fixed_bid')
    export_excel(output_df=output_asin_df,file_export_path='text.xlsx',sheet_name='Sheet1')