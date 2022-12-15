import pandas as pd
import openpyxl
import xlsxwriter
from datetime import date

def create_single_row(columns:list,cam_id):
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


def createDataFrame(cam_id,budget,sku,bid_list:list,billing_strategy,date,tactic,targeting_list:list,match_type,targeting_type):
    columns = ['Product','Entity','Operation','Campaign Id','Ad Group Id','Campaign Name','Ad Group Name','Start Date','End Date','State','Tactic','Budget Type','Budget','Cost Type','Bid Optimization','SKU','ASIN','Ad Group Default Bid','Bid','Targeting Expression']
    df = pd.DataFrame(columns=columns)
    for i in range(0,3):
        df = df.append(create_single_row(columns=columns,cam_id=cam_id),ignore_index=True)
    
    df.iloc[0]['Entity'] = 'Campaign'
    df.iloc[0]['Campaign Name'] = cam_id
    df.iloc[0]['Start Date'] = date
    df.iloc[0]['Tactic'] = tactic
    df.iloc[0]['Budget Type'] = 'daily'
    df.iloc[0]['Budget'] = budget
    if targeting_type.lower()=='contextual vcpm' or targeting_type.lower()=='audience vcpm':
        df.iloc[0]['Cost Type'] = 'vCPM'
    elif targeting_type.lower()=='contextual cpc' or targeting_type.lower()=='audience cpc':
        df.iloc[0]['Cost Type'] = 'CPC'
    else:
        df.iloc[0]['Cost Type'] = ''

    df.iloc[1]['Entity'] = 'Ad group'
    df.iloc[1]['Ad Group Id'] = cam_id
    df.iloc[1]['Ad Group Name'] = cam_id
    df.iloc[1]['Ad Group Name'] = cam_id
    df.iloc[1]['Ad Group Default Bid'] = 1
    if match_type.lower() == 'reach':
        df.iloc[1]['Bid Optimization'] = 'reach'
    elif match_type.lower() =='page visits':
        df.iloc[1]['Bid Optimization'] = 'Optimize for page visits'
    elif match_type.lower() == 'cvr':
        df.iloc[1]['Bid Optimization'] = 'Optimize for conservations'
    else:
        df.iloc[1]['Bid Optimization'] = ''
    
    
    df.iloc[2]['Entity'] = 'Product ad'
    df.iloc[2]['Ad Group Id'] = cam_id
    df.iloc[2]['SKU'] = sku
    
    for i in range(len(targeting_list)):
        temp_row = create_single_row(columns=columns,cam_id=cam_id)
        temp_row['Entity'] = 'Product targeting'
        temp_row['Ad Group Id'] = cam_id
        if len(bid_list)>1:
            if i < len(bid_list):
                temp_row['Bid'] = bid_list[i]
            else:
                temp_row['Bid'] = bid_list[-1]
        else:
            temp_row['Bid'] = bid_list[0]
        if targeting_list[i].startswith("B0"):
            temp_row['Targeting Expression'] = '''asin="''' +str(targeting_list[i])+ '''"'''
        else:
            temp_row['Targeting Expression'] = str(targeting_list[i])
        df = df.append(temp_row,ignore_index=True)
    
    return df

def createResultDataFrame(input_df:pd.DataFrame):
    frameslist=[]
    for i in range(len(input_df)):
        cam_id = str(input_df.iloc[i]['CODE']) + ' ' + str(input_df.iloc[i]['Market']) +' ' +  str(input_df.iloc[i]['PPC Type']) + ' ' + str(input_df.iloc[i]['Tactic']) + ' ' + str(input_df.iloc[i]['Match Type']) +' ' +  str(input_df.iloc[i]['BRAND']) +' ' +  str(input_df.iloc[i]['Date'])+' ' + str(input_df.iloc[i]['PIC'])+' ' + str(input_df.iloc[i]['STT'])
        bid_list = str(input_df.iloc[i]['Bid']).split(',')
        targeting_list = str(input_df.iloc[i]['Targeting']).split(',')
        print(bid_list)
        print(targeting_list)
        temp_df = createDataFrame(cam_id = cam_id,budget=str(input_df.iloc[i]['Budget']),sku=str(input_df.iloc[i]['SKU']),bid_list=bid_list,
                                        billing_strategy=str(input_df.iloc[i]['Bid strategy']),date=str(input_df.iloc[i]['Date']),tactic=str(input_df.iloc[i]['Tactic']),
                                        targeting_list=targeting_list,match_type=str(input_df.iloc[i]['Match Type']),targeting_type=str(input_df.iloc[i]['Targeting type']))
        frameslist.append(temp_df)
    
    output_df = pd.concat(frameslist)
    return output_df


if __name__ == '__main__':
    input_df = pd.read_excel('../sample_files/Input_SD.xlsx',index_col=False)
    df_out = createResultDataFrame(input_df=input_df)
    print(df_out)