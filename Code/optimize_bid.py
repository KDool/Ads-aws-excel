import pandas as pd
import xlsxwriter
from datetime import date

#Read Bulk Product search term function
def read_bulk_report_SP(excel_file=''):
    bulk_df = pd.read_excel(excel_file,index_col=False)
    df_process = bulk_df[['Campaign Id','Ad Group Id','Portfolio Id','Keyword Id (Read only)','Campaign Name (Informational only)','Ad Group Name (Informational only)','Portfolio Name (Informational only)','Keyword Text','Campaign State (Informational only)']]
    df_process.columns = ['Campaign Id','Ad Group Id','Portfolio Id','Keyword Id','Campaign Name','Ad Group Name','Portfolio Name','Keyword Text','Campaign State']
    df_process = df_process[df_process['Campaign State']!='archived']
    return df_process

#Read Bulk Brand function
def read_bulk_report_Brands(excel_file=''):
    bulk_df = pd.read_excel(excel_file,index_col=False)
    df_process = bulk_df[['Campaign Id','Ad Group Id (Read only)','Keyword Id (Read only)','Campaign Name (Informational only)','Portfolio Name (Informational only)','Keyword Text','Campaign State (Informational only)']]
    df_process.columns = ['Campaign Id','Ad Group Id','Keyword Id','Campaign Name','Portfolio Name','Keyword Text','Campaign State']
    df_process['Ad Group Name'] = df_process['Campaign Name']
    df_process = df_process[df_process['Campaign State']!='archived']
    return df_process

#read Bulk Display function
# def read_bulk_report_SD(excel_file=''):
#     bulk_df = pd.read_excel(excel_file,index_col=False)
#     df_process = bulk_df[['Campaign Id','Ad Group Id','Keyword Id (Read only)','Campaign Name (Informational only)','Portfolio Name (Informational only)','Keyword Text']]
#     df_process.columns = ['Campaign Id','Ad Group Id','Keyword Id','Campaign Name','Portfolio Name','Keyword Text']
#     return df_process
def read_bulk_report_SD(excel_file=''):
    bulk_df = pd.read_excel(excel_file,index_col=False)
    df_process = bulk_df[['Campaign Id','Ad Group Id','Campaign Name (Informational only)','Ad Group Name (Informational only)','Targeting Expression','Campaign State (Informational only)']]
    df_process.columns = ['Campaign Id','Ad Group Id','Campaign Name','Ad Group Name','Targeting Expression','Campaign State']
    df_process = df_process[df_process['Campaign State']!='archived']
    # df_process = bulk_df
    return df_process


#Filter function for SP
def filter_SP_downbid(CSP_df: pd.DataFrame,acos,clicks,spend):
    filerted_df_1 = CSP_df[CSP_df['Total Advertising Cost of Sales (ACOS) ']>acos]
    filtered_df_2 = CSP_df[CSP_df['Clicks']>clicks][CSP_df['7 Day Total Sales ']==0]
    filtered_df_3 = CSP_df[CSP_df['Spend']>spend][CSP_df['7 Day Total Sales ']==0]
    list_frames = [filerted_df_1,filtered_df_2,filtered_df_3]
    filtered_df = pd.concat(list_frames)
    filtered_df.drop_duplicates(inplace=True)
    return filtered_df


#Filter Function for SB
def filter_SB_downbid(CSP_df: pd.DataFrame,acos,clicks,spend):
    filerted_df_1 = CSP_df[CSP_df['Total Advertising Cost of Sales (ACOS) ']>acos]
    filtered_df_2 = CSP_df[CSP_df['Clicks']>clicks][CSP_df['14 Day Total Sales ']==0]
    filtered_df_3 = CSP_df[CSP_df['Spend']>spend][CSP_df['14 Day Total Sales ']==0]
    list_frames = [filerted_df_1,filtered_df_2,filtered_df_3]
    filtered_df = pd.concat(list_frames)
    filtered_df.drop_duplicates(inplace=True)
    return filtered_df

# Filter SD function
def filter_SD_downbid(CSP_df: pd.DataFrame,acos,spend):
    filerted_df_1 = CSP_df[CSP_df['Total Advertising Cost of Sales (ACOS) ']>acos]
    filtered_df_3 = CSP_df[CSP_df['Spend']>spend][CSP_df['14 Day Total Sales ']==0]
    list_frames = [filerted_df_1,filtered_df_3]
    filtered_df = pd.concat(list_frames)
    filtered_df.drop_duplicates(inplace=True)
    return filtered_df


# Filter Tang bid
def filter_upbid(CSP_df: pd.DataFrame,dates_diff,impressions):
    def filter_row_date(row):
        if (date.today() - row['Start Date'].date()).days > dates_diff:
            return True
        else:
            return False
    m = CSP_df.apply(filter_row_date,axis=1)
    filtered_df_1 = CSP_df[m]
    if '14 Day Total Sales ' in CSP_df.columns.to_list():
        filtered_df_2 = CSP_df[CSP_df['Impressions'] <impressions][CSP_df['14 Day Total Sales ']==0]
        list_frames = [filtered_df_1,filtered_df_2]
    elif '7 Day Total Sales ' in CSP_df.columns.to_list():
        filtered_df_2 = CSP_df[CSP_df['Impressions'] <impressions][CSP_df['7 Day Total Sales ']==0]
        list_frames = [filtered_df_1,filtered_df_2]
    else:
        list_frames = [filtered_df_1]
    
    filtered_df = pd.concat(list_frames)
    filtered_df.drop_duplicates(inplace=True)
    return filtered_df



def get_campid_toDF(filtered_df:pd.DataFrame,bulk_df:pd.DataFrame):

    bulk_df_campaign = get_campaign_infor(bulk_df)
    bulk_df_adgroup = get_adgroup_infor(bulk_df)

    filtered_df['Campaign Name'] = filtered_df['Campaign Name'].astype("string")
    bulk_df_campaign['Campaign Name'] = bulk_df_campaign['Campaign Name'].astype("string")
    df_ver1 = pd.merge(filtered_df, bulk_df_campaign, how='left',left_on = 'Campaign Name',right_on = 'Campaign Name')
    df_ver2 = pd.merge(df_ver1, bulk_df_adgroup, how='left',left_on = 'Ad Group Name',right_on = 'Ad Group Name')
    print('Length original: ',len(filtered_df))
    print('Length camp merge: ',len(df_ver1))
    print('Lengh adgroup merge: ',len(df_ver2))
    return df_ver2

def get_campaign_infor(df_bulk:pd.DataFrame):
    df_campaign_info = df_bulk[['Campaign Id','Campaign Name']]
    df_campaign_info.drop_duplicates(inplace=True)
    df_campaign_info.reset_index(inplace=True)
    return df_campaign_info

def get_adgroup_infor(df_bulk:pd.DataFrame):
    df_adgroup_info = df_bulk[['Ad Group Id','Ad Group Name']]
    df_adgroup_info.drop_duplicates(inplace=True)
    df_adgroup_info.reset_index(inplace=True)
    return df_adgroup_info


def create_row_optimizeBid_asin(columns:list,camp_id,ad_group_id,bid,increase_number):
    row = {}
    for item in columns:
        if item == 'Entity':
            row[item] = 'Product Targeting'
        elif item == 'Product':
            row[item] = 'Sponsored Products'
        elif item =='Operation':
            row[item] = 'Update'
        elif item == 'Campaign_Id':
            row[item] = camp_id
        elif item == 'Ad_Group_Id':
            row[item] = ad_group_id
        elif item == 'Bid':
            row[item] = bid
        elif item == 'Ad_Group_Default_Bid':
            row[item] = str(float(bid)*increase_number)
        else:
            row[item] = None
    return row

def create_row_optimizeBid_keyword(columns:list,camp_id,ad_group_id,bid,increase_number):
    row = {}
    for item in columns:
        if item == 'Entity':
            row[item] = 'Keyword'
        elif item == 'Product':
            row[item] = 'Sponsored Products'
        elif item =='Operation':
            row[item] = 'Update'
        elif item == 'Campaign_Id':
            row[item] = camp_id
        elif item == 'Ad_Group_Id':
            row[item] = ad_group_id
        elif item == 'Bid':
            row[item] = bid
        elif item == 'Ad_Group_Default_Bid':
            row[item] = str(float(bid)*increase_number)
        else:
            row[item] = None
    return row





def export_excel_files(filtered_df:pd.DataFrame,old_bid,increase_number):
    columns = ['Product','Entity','Operation','Campaign_Id', 'Ad_Group_Id',	'Portfolio_Id','Ad_Id_Read_only','Keyword_Id_Read_only','Product_Targeting_Id_Read_only',	
            'Campaign_Name_Informational_only',	'Ad_Group_Name_Informational_only',	'Start_Date',	'End_Date',	'Targeting_Type',	'State',	'Daily_Budget',	
            'SKU',	'ASIN_Informational_only',	'Bid',	'Ad_Group_Default_Bid',	'Keyword_Text',	'Match_Type',	'Bidding_Strategy',	'Placement',	'Percentage',
            'Product_Targeting_Expression']
    result_df = pd.DataFrame(columns=columns)
    for i in range(len(filtered_df)):
        if filtered_df.iloc[i]['Targeting'].startswith('asin'):
            temp_asin_row = create_row_optimizeBid_asin(columns = columns,camp_id=filtered_df.iloc[i]['Campaign Id'],ad_group_id=filtered_df.iloc[i]['Ad Group Id'],bid=old_bid,increase_number=increase_number)
            result_df = result_df.append(temp_asin_row,ignore_index=True)
        else:
            temp_kw_row = create_row_optimizeBid_keyword(columns = columns,camp_id=filtered_df.iloc[i]['Campaign Id'],ad_group_id=filtered_df.iloc[i]['Ad Group Id'],bid=old_bid,increase_number=increase_number)
            result_df = result_df.append(temp_kw_row,ignore_index=True)         
    return result_df


# if __name__ == '__main__':
#     df_sp_targeting = pd.read_excel('../sample_files/Sponsored Products Targeting report .xlsx',index_col=False)
#     df_sd_targeting = pd.read_excel('../sample_files/SD Targeting report.xlsx',index_col=False)
#     df_sb_keyword = pd.read_excel('../sample_files/Sponsored Brands Keyword report .xlsx')
#     # print(df_sp_targeting.head(5))
#     # print(df_sd_targeting.head(5))
#     # print(df_sb_keyword.head(5))
#     # print('LEN: ',len(df_sb_keyword))
#     # print('LEN: ',len(df_sd_targeting))
#     # sb_filter_downbid = filter_SB_downbid(df_sb_keyword,acos=0.25,clicks=8,spend=10)
#     # sp_filter_downbid = filter_SP_downbid(df_sp_targeting,acos=0.25,clicks=8,spend=10)
#     # sd_filter_downbid = filter_SD_downbid(df_sd_targeting,acos=0.25,spend=10)

#     sp_filter_upbid = filter_upbid(df_sp_targeting,dates_diff=35,impressions=5)
#     sb_filter_upbid = filter_upbid(df_sb_keyword,dates_diff=35,impressions=5)
#     sd_filter_upbid = filter_upbid(df_sd_targeting,dates_diff=35,impressions=5)
#     print(sb_filter_upbid.columns.to_list())
#     print(sp_filter_upbid.columns.to_list())
#     print(sd_filter_upbid.columns.to_list())