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
        else:
            single_row_dictionary[item] = None
    return single_row_dictionary

def createDataFrame_Auto(cam_id,budget,productName,bid,bidding_strategy,percentage,date):
    columns=['Product','Entity','Operation','Campaign Id','Ad Group Id','Portfolio Id',
                           'Ad Id','Keyword Id','Product Targeting Id','Campaign Name','Ad Group Name',
                           'Start Date','End Date','Targeting Type','State','Daily Budget','SKU','Asin',
                           'Ad Group Default Bid','Bid','Keyword Text','Match Type','Bidding Strategy',
                           'Placement','Percentage','Product Targeting Expression']
    df = pd.DataFrame(columns=columns)
    # 5 dong co dinh
    for i in range (0,5):
        df = df.append(create_single_row_SP(columns=columns,cam_id=cam_id),ignore_index=True)
        
    df.iloc[0]['Entity'] = 'Campaign'
    df.iloc[0]['Campaign Name'] = cam_id
    df.iloc[0]['Targeting Type'] = 'Auto'
    df.iloc[0]['Start Date'] = str(date)
    df.iloc[0]['Daily Budget'] = budget
    df.iloc[0]['Bidding Strategy'] = bidding_strategy
    df.iloc[0]['State'] = 'Enable'


    df.iloc[1]['Entity'] = 'Bidding Adjustment'
    df.iloc[1]['Placement'] = 'placementTop'
    df.iloc[1]['Percentage'] = percentage

    df.iloc[2]['Entity'] = 'Bidding Adjustment'
    df.iloc[2]['Placement'] = 'placementProductPage'
    df.iloc[2]['Percentage'] = percentage

    df.iloc[3]['Entity'] = 'Ad group'
    df.iloc[3]['Ad Group Id'] = cam_id
    df.iloc[3]['Ad Group Name'] = cam_id
    df.iloc[3]['State'] = 'Enable'
    df.iloc[3]['Ad Group Default Bid'] = productName
    # df.iloc[3]['Ad Group Default Bid'] = bid
    
    df.iloc[4]['Entity'] = 'Product ad'
    df.iloc[4]['Ad Group Id'] = cam_id
    df.iloc[4]['SKU'] = productName
    df.iloc[4]['State'] = 'Enable'

    return df

def createResultDataFrame(input_df:pd.DataFrame):
    framelist = []
    for i in range(len(input_df)):
        cam_id = str(input_df.iloc[i]['CODE']) +  str(input_df.iloc[i]['Market']) + str(input_df.iloc[i]['PPC Type']) + str(input_df.iloc[i]['Match type'])+ str(input_df.iloc[i]['BRAND']) + str(input_df.iloc[i]['Date']) +str(input_df.iloc[i]['PIC'] + str(input_df.iloc[i]['STT']))
        temp_df = createDataFrame_Auto(cam_id=cam_id,budget=input_df.iloc[i]['Budget'],productName=input_df.iloc[i]['SKU'],
                                        bid=input_df.iloc[i]['Bid'], bidding_strategy=input_df.iloc[i]['Bid strategy'],percentage="{:.0%}".format(input_df.iloc[i]['Percentage']),date=str(input_df.iloc[i]['Date']))
        
        framelist.append(temp_df)
    output_df = pd.concat(framelist)
    return output_df


if __name__ == '__main__':
    input_df = pd.read_excel('../sample_files/Input_SP_Auto.xlsx',index_col=False)
    df_out = createResultDataFrame(input_df=input_df)
    print(df_out)