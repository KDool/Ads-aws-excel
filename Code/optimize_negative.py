import pandas as pd
import xlsxwriter
from datetime import date
from sqlalchemy import create_engine


def read_CSP_report(excel_file=''):
    report_df = pd.read_excel(excel_file,index_col=False)
    return report_df

def read_bulk_report_CSP(excel_file=''):
    bulk_df = pd.read_excel(excel_file,index_col=False)
    df_process = bulk_df[['Campaign Id','Ad Group Id','Portfolio Id','Keyword Id (Read only)','Campaign Name (Informational only)','Portfolio Name (Informational only)','Keyword Text','Campaign State (Informational only)']]
    df_process.columns = ['Campaign Id','Ad Group Id','Portfolio Id','Keyword Id','Campaign Name','Portfolio Name','Keyword Text','Campaign State']
    df_process = df_process[df_process['Campaign State']!='archived']
    return df_process

def read_bulk_report_Brands(excel_file=''):
    bulk_df = pd.read_excel(excel_file,index_col=False)
    df_process = bulk_df[['Campaign Id','Ad Group Id (Read only)','Keyword Id (Read only)','Campaign Name (Informational only)','Portfolio Name (Informational only)','Keyword Text','Campaign State (Informational only)']]
    df_process.columns = ['Campaign Id','Ad Group Id','Keyword Id','Campaign Name','Portfolio Name','Keyword Text','Campaign State']
    df_process = df_process[df_process['Campaign State']!='archived']
    return df_process


# Rules Filter Function
def filter_CSP_negative(CSP_df: pd.DataFrame,acos,clicks,spend):
    filerted_df_1 = CSP_df[CSP_df['Total Advertising Cost of Sales (ACOS) ']>acos]
    filtered_df_2 = CSP_df[CSP_df['Clicks']>clicks][CSP_df['7 Day Total Sales ']==0]
    filtered_df_3 = CSP_df[CSP_df['Spend']>spend][CSP_df['7 Day Total Sales ']==0]
    list_frames = [filerted_df_1,filtered_df_2,filtered_df_3]
    filtered_df = pd.concat(list_frames)
    filtered_df.drop_duplicates(inplace=True)
    return filtered_df

def filter_Brands_negative(CSP_df: pd.DataFrame,acos,clicks,spend):
    filerted_df_1 = CSP_df[CSP_df['Total Advertising Cost of Sales (ACOS) ']>acos]
    filtered_df_2 = CSP_df[CSP_df['Clicks']>clicks][CSP_df['14 Day Total Sales ']==0]
    filtered_df_3 = CSP_df[CSP_df['Spend']>spend][CSP_df['14 Day Total Sales ']==0]
    list_frames = [filerted_df_1,filtered_df_2,filtered_df_3]
    filtered_df = pd.concat(list_frames)
    filtered_df.drop_duplicates(inplace=True)
    return filtered_df

# GET Campaign ID related to Campagin Name columns from DB - table_bulk_products
# Return Dataframe with Campaign Id columns
def get_campid_toDF(filtered_df:pd.DataFrame,bulk_df=pd.DataFrame):
    filtered_df['Campaign Name'] = filtered_df['Campaign Name'].astype("string")
    bulk_df['Campaign Name'] = bulk_df['Campaign Name'].astype("string")
    df_cd = pd.merge(filtered_df, bulk_df, how='left',left_on = 'Campaign Name',right_on = 'Campaign Name')

    return df_cd


def create_row_dictonary_kw(columns:list,Targeting='',campaign_id='',ad_group_id=''):
    row_dictionary = {}
    for item in columns:
        if item =='Product':
            row_dictionary[item] = 'Sponsored products'
        elif item == 'Entity':
            row_dictionary[item] = 'Negative keyword'
        elif item == 'Operation':
            row_dictionary[item] = 'Create'
        elif item == 'Campaign Id':
            row_dictionary[item] = campaign_id
        elif item == 'Ad Group Id':
            row_dictionary[item] = ad_group_id
        elif item == 'State':
            row_dictionary[item] = 'enabled'
        elif item == 'Keyword Text':
            row_dictionary[item] = Targeting
        elif item == 'Match Type':
            row_dictionary[item] = 'negativeExact'
        else:
            row_dictionary[item] = None
    return row_dictionary

def create_row_dictonary_asin(columns:list,Targeting='',campaign_id='',ad_group_id=''):
    row_dictionary = {}
    for item in columns:
        if item =='Product':
            row_dictionary[item] = 'Sponsored products'
        elif item == 'Entity':
            row_dictionary[item] = 'Negative keyword'
        elif item == 'Operation':
            row_dictionary[item] = 'Create'
        elif item == 'Campaign Id':
            row_dictionary[item] = campaign_id
        elif item == 'Ad Group Id':
            row_dictionary[item] = ad_group_id
        elif item == 'State':
            row_dictionary[item] = 'enabled'
        elif item == 'Product Targeting Expression':
            row_dictionary[item] = Targeting
        elif item == 'Match Type':
            row_dictionary[item] = 'negativeExact'
        else:
            row_dictionary[item] = None
    return row_dictionary


def export_excel_files(filtered_df:pd.DataFrame):
    columns=['Product','Entity','Operation','Campaign Id','Ad Group Id','Portfolio Id',
                           'Ad Id','Keyword Id','Product Targeting Id','Campaign Name','Ad Group Name',
                           'Start Date','End Date','Targeting Type','State','Daily Budget','sku','asin',
                           'Ad Group Default Bid','Bid','Keyword Text','Match Type','Bidding Strategy',
                           'Placement','Percentage','Product Targeting Expression']
    result_df = pd.DataFrame(columns=columns)
    for i in range(len(filtered_df)):
        if filtered_df.iloc[i]['Targeting'].startswith('asin'):
            temp_asin_row = create_row_dictonary_asin(columns = columns,Targeting = filtered_df.iloc[i]['Targeting'],campaign_id=filtered_df.iloc[i]['Campaign Id'],ad_group_id=filtered_df.iloc[i]['Ad Group Name'])
            result_df = result_df.append(temp_asin_row,ignore_index=True)
        else:
            temp_kw_row = create_row_dictonary_kw(columns = columns,Targeting = filtered_df.iloc[i]['Targeting'],campaign_id=filtered_df.iloc[i]['Campaign Id'],ad_group_id=filtered_df.iloc[i]['Ad Group Name'])
            result_df = result_df.append(temp_kw_row,ignore_index=True)         
    return result_df


if __name__ == '__main__':
    df_CSP = read_CSP_report('../sample_files/Sponsored Products Search term report .xlsx')
    df_bulk = read_bulk_report('../sample_files/BULK Sponsored Products Campaigns.xlsx')
    df_CSP_filter = filter_CSP_negative(CSP_df=df_CSP,acos=0.6,clicks=20,spend=20)
    # df_x = get_campid_toDF('table_bulk_products')
    df_campaign_info = df_bulk[['Campaign Id','Campaign Name']]
    df_campaign_info.drop_duplicates(inplace=True)
    df_campaign_info.reset_index(inplace=True)

    # df_campaign_info.drop_duplicates(inplace=True,subset=['Campaign Id','Campaign Name'])
    
    df_cd = get_campid_toDF(filtered_df=df_CSP_filter,bulk_df=df_campaign_info)

    df_export = export_excel_files(filtered_df=df_cd)
    print(df_export)
    writer = pd.ExcelWriter("Optimize.xlsx", engine='xlsxwriter')
    df_export.to_excel(writer,sheet_name = 'Sheet1', index=False)
    # output_keyword_df.to_excel(writer,sheet_name = 'sheet_keyword', index=False)
    writer.save() 
