import pandas as pd
import numpy as np
import datetime

import os
import glob
from shutil import copyfile
from vincent.colors import brews
import math
import sys
import xlsxwriter

########################################################## Input Section ###############################################
## Configuration file path with name

error = 'Error in Date.xlsx'

dir_path = ''
output_path = ''
project_dict = {}
start_date = ''
end_date = ''
summary_file = ''
review_file = ''
review_sh = ''
service_file = ''
service_sh = 'Sheet1'
##Feedback_report_sheets
sheets = []

## Summary report sheets
summary_sh = 'Summary'
category_sh = 'All categories'
project_sh = 'All projects'

########################################################## Input End ###############################################

## summary report information
validation_sh = 'Data Validations'
metric_tb = 'Findings Metrics'
metric_head = ('Artifact Name', 'Total # of QA', '# of Open Findings','# of Closed Findings','# of Deferred Findings')
type_tb = 'Artifact Finding Types and Count'
left_tb1 = ['# of Test Cases','# of BR\'s', '# of FR\'s', '# of Test Results']
left_tb2 = ['Vendor Name']

## Summary: project sheet
projects_headers = []
projects_list = []
project_path = ''

## Data Validations: "finding types", the name of the type and the position in excel
types_dict = {'Test Plan':'A:1:20','Test Cases':'E:1:21' ,'RTM':'C:1:10',\
              'Test Results':'C:13:25', 'Test Summary Report':'C:30:33'}
##              'Requirements':'E:26:41'}
headers_dict = {}
processed_data_dict = {}
finding_type_df = []

find_types_df = []
find_metrics_df = []
find_metric_list = []

count_testcase = 0
count_testresult = 0

sum_tc = 0
sum_testcase = 0
sum_tr = 0
summary_df = pd.DataFrame()
summary_findtypes_dict = {}
##project_num = 0
project_end = False

error_review = []
error_followup1 = []
error_followup2 = []

## feedback review report
total_opened = 0
total_closed = 0
total_remain = 0
review_list =[]
list_inequal = []
review_open_df= pd.DataFrame()
review_header = ['S.No','Project ID','Portfolio','Project Name','PM','Project Phase','Project Stage','QA Reviews Status',\
                  'Date of Production','QA resource','Total # of TC\'s Reviewed','Total # of TR\'s Reviewed']
review_sub = ['# opened', '# closed','# remained open']
review_cols = []

## Columns need to be analyzed
col_type = 'Finding Type'
col_status = 'Resolution Status\n (Open / Closed)'
col_severity = 'Severity'
col_testcase_count = 'Test Case Count'

## Write summary texts for each project
service_tc_dict = {}
service_tr_dict = {}


'''
    v1.6: do not count tracking issues and clarification requests

'''
tracking_str_list = ['failed test cases - tracking purpose','clarification request', 'test execution in progress - tracking  purpose']

def analyze_tracking_clarify_observations(artifact, dataframe):
    
    #Tracking  Purpose and Clarification Request
    tracking_df = pd.DataFrame()
    tracking_df = dataframe.loc[dataframe[col_type].str.lower().isin(tracking_str_list)]

    return len(tracking_df), len(tracking_df.loc[tracking_df[col_status] == 'Open']), len(tracking_df.loc[tracking_df[col_status] == 'Closed']), len(tracking_df.loc[tracking_df[col_status] == 'Deferred'])    



## copy source files from source paths:project_dict
def get_source(proj_path_dict, copy_path):
    print(proj_path_dict)
    print(copy_path)
    for project in proj_path_dict:
        ## path of the project
        path = proj_path_dict[project]
        print(path)

        ## get all file of the folder
        list_of_files = glob.glob(path + r'\*')
        latest_file = max(list_of_files, key=os.path.getmtime)

        file_name = os.path.basename(latest_file)
        copy_file = os.path.join(copy_path, file_name)
        ## Add new path to the project
        project_dict[project] = copy_file
        ## copy the source file to the new folder
        copyfile(latest_file, copy_file)

    print("Complete copying files")
    
def get_source(proj_path_dict):
    for project in proj_path_dict:
        ## path of the project
        path = proj_path_dict[project]

        ## get all file of the folder
        list_of_files = glob.glob(path + r'\\\*')
        latest_file = max(list_of_files, key=os.path.getmtime)

        file_name = os.path.basename(latest_file)
        copy_file = os.path.join(copy_path, file_name)
        ## Add new path to the project
        project_dict[project] = copy_file
        ## copy the source file to the new folder
        print("TTTTTTTTTTTTTTTTTTTTTTTTTT\n")
        print(latest_file)
        print(copy_file)
        copyfile(latest_file, copy_file)

    print("Complete copying files")
    

## Read configuration files before run the program\
def read_files_software(project_dict):
    print(start_date,end_date,summary_file,review_file,review_sh,sheets,review_cols,projects_headers,project_path,service_file,copy_path)
    get_source(project_dict)
    writer_summary = pd.ExcelWriter(summary_file,engine='xlsxwriter')


## Summary: Read and analyze data of finding types 
def get_valid_type(project):
    ## read and analyze data of finding types
    for item in types_dict:
        print("---ITEM---",item)
        find_type = types_dict[item]
        pos = find_type.split(':')

        ## Get data with specific finding type in validation sheet
        type_df = pd.read_excel(project_dict[project],validation_sh, usecols = pos[0], skip_blank_lines=False)
        print('@@@@@@@@@@@@@@@@@@@@@@@@@@@get_valid_type')


        i = int(pos[1]) - 1 
        j = int(pos[2]) - 1
        ## Select the list of finding type
        type_df = type_df.iloc[i:j]
        type_df.rename(columns={list(type_df)[0]:item+' Finding Type'}, inplace=True)

        ## Add 'count' column
        type_df['count'] = [0] * len(type_df)

        ## Get data from sheet
        counts = analyze_data(project, item)
        # print('counts:')
        # print(counts)

        # print(type_df)

        if counts is not None:
            ## Input data in 'count' column
            type_df['count'] = type_df[item+' Finding Type'].map(counts)
            type_df['count'] = type_df['count'].fillna(0)
            type_df['count'] = type_df['count'].astype(int)

            ## Add 'Grand Total' row
            sum_count = type_df['count'].astype(int).sum()
            grand_total = ['Grand Total', sum_count]
            # print(type_df)
            # type_df.loc[len(type_df)]= grand_total
            total_df = pd.DataFrame([grand_total], columns=[item+' Finding Type','count'])
            # print(total_df)
            type_df = type_df.append(total_df, ignore_index=True)


            # print(type_df)
        else:
            grand_total = ['Grand Total', 0]
            # type_df.loc[len(type_df)]= grand_total

            total_df = pd.DataFrame([grand_total], columns=[item+' Finding Type','count'])
            # print(total_df)
            type_df = type_df.append(total_df, ignore_index=True)


        # print(type_df)
        
        ## Combine all data from finding types per project
        find_types_df.append(type_df)

        ## summary
        if summary_findtypes_dict[item +' Finding Type'].empty:
            summary_findtypes_dict[item+ ' Finding Type'] = type_df
        else:
            summary_findtypes_dict[item+ ' Finding Type']['count'] = summary_findtypes_dict[item+ ' Finding Type']['count'].add(type_df['count'],\
                                                                                                                                fill_value=0)
            

## Summary: Analyze the data in each sheet and get the result
## Summary: Analyze the data in each sheet and get the result
def analyze_data(project, sheet):
    global total_opened, total_closed, total_remain
    global count_testcase, count_testresult, list_inequal
    global review_list, error_review, error_followup1, error_followup2
    global sum_tc, sum_tr, sum_testcase
    print(project_dict[project],'****'+sheet)
    
    ## Get data with specific finding type in validation sheet
    sheet_df = pd.read_excel(project_dict[project],sheet, na_values=['NA'])

    # Get index of tables
    sheet_df.reset_index()
    idx_list = sheet_df[sheet_df['CITSQ- QA Feedback Report']=='S.No'].index.tolist()
    idx = 7
    if(idx_list):
        print("True")
        idx = idx_list[0]
    else:
        print("IDX LIST IS empty")
        idx = 7

    head = idx - 1

    header_df = sheet_df.iloc[0:head,:]
    print("\nHEADER_DF \n", header_df)
    sheet_df = sheet_df.iloc[idx:,:]
    print("\nSHEET_DF \n", sheet_df)


    ## Replace the header
    header = sheet_df.iloc[0]
    print("\nheader \n", header)

    sheet_df.columns = header.tolist()
    print("\nsheet_df.columns \n", sheet_df.columns)

    sheet_df = sheet_df[1:]

    ####################################
    ## Get data from the time period
    sheet_df = extract_data(project,sheet,sheet_df)
    processed_data_dict[sheet] = sheet_df
    headers_dict[sheet] = header_df


    if sheet_df.empty:
        find_metric_list.append([sheet,0,0,0,0])
        list_sheet_count = [0,0,0]
        review_list += list_sheet_count
        return None


    # Remove finding type:tracking issue, clarification
    track_total, track_open, track_closed, track_deferred = analyze_tracking_clarify_observations(sheet, sheet_df)

    # count the values
    counts_type = sheet_df.groupby(col_type).size()
##    print(counts_type.keys().tolist())

    # Finding_type_df filter the nan value
    valid_findings = counts_type.keys().tolist()
    finding_type_df = sheet_df.loc[sheet_df[col_type].isin(valid_findings)]
    counts_status = finding_type_df.groupby(col_status).size()


    # Finding type filter for count column
    count_df = sheet_df.loc[sheet_df[col_type].isin(tracking_str_list)]
    track_count = 0
    if sheet == 'Test Cases' or sheet == 'Test Results':
        track_count = count_df[col_testcase_count].sum()*1
##        print('$$$$$$$$$$$$$$$$track_count:' + str(track_count))
    

    # sum of test cases count
    if sheet == 'Test Cases':
        count_testcase = sheet_df[col_testcase_count].sum() * 1 # if boolean, turn to int
        count_testcase = count_testcase - track_count
        sum_tc += count_testcase
        sum_testcase += count_testcase
##        print('Total # of TC\'s Reviewed: ' + str(count_testcase))

    ## sum of test result count
    if sheet == 'Test Results':
        count_testresult = sheet_df[col_testcase_count].sum() * 1
        count_testresult = count_testresult - track_count
        sum_tr += count_testresult
##        print('Total # of TR\'s Reviewed: '+ str(count_testresult))



    ## get row of finding metrics
    types = finding_type_df[col_type]
    types = types[types.notnull()]
    type_digit = sheet_df[col_type].astype(str).str.isdigit().sum()

    count_total = len(types)-track_total
    name = sheet

    print('#############################Tracking')
    print('total: ' + str(track_total))
    print('Open: ' + str(track_open))
    print('Closed: ' + str(track_closed))
    print('Referred: ' + str(track_deferred))

    count_open = counts_status.get('Open')
##    count_open = len(count_type_status.loc[count_type_status[col_status] == 'open'])
    count_open = count_open-track_open if count_open is not None else 0
    count_closed = counts_status.get('Closed')
    count_closed = count_closed-track_closed if count_closed is not None else 0    
    count_deferred = counts_status.get('Deferred')
    count_deferred = count_deferred-track_deferred if count_deferred is not None else 0
    print('#############################Finding metrics')
##    print(counts_status)
    # print(str(count_total))
    # print(str(count_open))
    # print(str(count_closed))
    # print(str(count_deferred))
    
    list_row = [name,count_total,count_open,count_closed,count_deferred]
    list_row = [0 if x == None else x for x in list_row ]
    find_metric_list.append(list_row)


    ###################################################################################################
    ############## get values of feedback review report
    
    cod_open = (review_open_df[col_status] == 'Open') | (review_open_df[col_status] == 'Deferred')\
               | ((review_open_df[col_status] == 'Closed'))

    ## Get rid of tracking and clarification categories
    review_opened = review_open_df.loc[~review_open_df[col_type].str.lower().isin(tracking_str_list)]
    count_review_open = len(review_opened[cod_open])

    ## closed count
    cod_close = (sheet_df[col_status] == 'Closed')
    count_review_close = len(sheet_df[cod_close])-track_closed
    ## remain open count
    cod_remain = (sheet_df[col_status] == 'Open') | (sheet_df[col_status] == 'Deferred')
    count_review_remain = len(sheet_df[cod_remain])-track_deferred-track_open
    ## Sum
    total_opened += count_review_open
    total_closed += count_review_close
    total_remain += count_review_remain
    ## create a list for these count, order:opened, closed, remained
    list_sheet_count = [count_review_open, count_review_close, count_review_remain]
    review_list += list_sheet_count

    ## review services report
    if sheet == 'Test Cases' or sheet == 'Test Results':
        # review_service_df = sheet_df.loc[~sheet_df[col_type].str.lower().isin(tracking_str_list)]
        
        sum_tc = sheet_df[col_testcase_count].sum()
        opened_df = review_open_df[cod_open]
        obs_ser = opened_df.groupby(col_type).size()
        obs_list = obs_ser.index.tolist()
        tracking_open = [x for x in obs_list if x.lower() in tracking_str_list]
        obs_list = [x for x in obs_list if x.lower() not in tracking_str_list]
        
        # Remove tracking
        open_df = opened_df.loc[~sheet_df[col_type].str.lower().isin(tracking_str_list)]

        
        write_service(project,sheet,sum_tc,len(open_df),obs_list, count_review_close)


    return counts_type

    
## Summary: extract data from a specific time period
def extract_data(project,sheet,sheet_df):
    global writer_process, review_open_df

    ## Columns need to be analyzed
    col_review_date = 'Review Date'
    col_followup_date1 = 'Follow up Review Date 1'
##    if project == 'iPortal 1.5 Release 3' and sheet == 'Test Cases':
##        col_followup_date2 = 'Follow up Review Date 2&3'
##    else:
    col_followup_date2 = 'Follow up Review Date 2'

    #######################################################################################
    ## Get data from the time range
    ## Codition on the time period:review date, followup review date1, followup review date2
##    print(sheet_df)
    review_date = sheet_df[col_review_date]
    followup_date1 = sheet_df[col_followup_date1]
    followup_date2 = sheet_df[col_followup_date2]

    ## Debugging on which rows in a certain column contains str format instead of datetime
##    print('String in Review Date')
##    print(sheet_df[sheet_df[col_review_date].apply(lambda x: type(x)==str)])
##    print('String in Follow up Review Date 1')
##    print(sheet_df[sheet_df[col_followup_date1].apply(lambda x: type(x)==str)])
##    print('String in Follow up Review Date 2')
##    print(sheet_df[sheet_df[col_followup_date2].apply(lambda x: type(x)==str)])
    ## Write rows which contains string into error excel
    error_review = sheet_df[sheet_df[col_review_date].apply(lambda x: type(x)==str)]
    error_followup1 = sheet_df[sheet_df[col_followup_date1].apply(lambda x: type(x)==str)]
    error_followup2 = sheet_df[sheet_df[col_followup_date2].apply(lambda x: type(x)==str)]
    writer = pd.ExcelWriter(error,engine='xlsxwriter')
    error_review.to_excel(writer, sheet_name = 'Review Date', startrow=0, startcol=0, index=False, header=False)
    error_followup1.to_excel(writer, sheet_name = 'Followup Date 1', startrow=0, startcol=0, index=False, header=False)
    error_followup2.to_excel(writer, sheet_name = 'Followup Date 2', startrow=0, startcol=0, index=False, header=False)
    writer.save()


    review_time_range = (review_date >= start_date) & (review_date <= end_date)
    followup_time_range1 = (followup_date1 >= start_date) & (followup_date1 <= end_date)
    followup_time_range2 = (followup_date2 >= start_date) & (followup_date2 <= end_date)
    df_time = followup_time_range2 | followup_time_range1 | review_time_range
    sheet_df = sheet_df[df_time]
    print(review_time_range)
    print("\nSHEEEEEETttt\n")
    print(sheet_df[review_time_range])
    review_open_df = sheet_df[review_time_range]
    
    return sheet_df


## Write the data within time range into excel
def write_processed_data(project):
##    project_name = project if len(project) <= 31 else project[:30]
    
##    writer_process = pd.ExcelWriter(project_path + '//'+project_name+'_processed.xlsx',engine='xlsxwriter')
    
    writer_process = pd.ExcelWriter(project_path + '\\'+project+'_processed.xlsx',engine='xlsxwriter')
    for sheet in processed_data_dict:
        headers_dict[sheet].to_excel(writer_process, sheet_name = sheet, startrow=0, index=False)
        processed_data_dict[sheet].to_excel(writer_process, sheet_name = sheet, startrow=7, index=False)
        
    writer_process.save()
    
'''
    review_service report
'''
def write_service(project,sheet,sum_tc,sum_opened,obs_list,sum_closed):
    
    #line1 = 'Review of ' + str(sum_tc) + ' ' + sheet
    line1 = 'Review of ' + str(sum_tc)
    line2 = str(sum_opened) + ' observations opened due to ' + ", ".join(obs_list)
    line3 = str(sum_closed) + ' observations closed'
    #task = line1 + '; ' + line2 + '; ' + line3

    if sheet == 'Test Cases':
        line1 = line1 + ' ' + sheet
        task = line1 + '; ' + line2 + '; ' + line3
        service_tc_dict[project] = task
    else:
        line1 = line1 + ' Test Evidences'
        task = line1 + '; ' + line2 + '; ' + line3
        service_tr_dict[project] = task
        
    header = ['Project', 'Test Cases', 'Test Evidences']
    if project_end:
        service_list = []
        ## combine test cases and test evidences
        for key in project_dict:
            if key in service_tc_dict:
                if key in service_tr_dict:
                    service_list.append([key,service_tc_dict[key],service_tr_dict[key]])
                else:
                    service_list.append([key,service_tc_dict[key],np.nan])
            elif key in service_tr_dict:
                service_list.append([key,np.nan,service_tr_dict[key]])
            else:
                service_list.append([key,np.nan,np.nan])
            
        df = pd.DataFrame(service_list)
        df.columns = header
        writer_service = pd.ExcelWriter(service_file,engine='xlsxwriter')
        df.to_excel(writer_service, sheet_name = service_sh,index=False)
        writer_service.save()

## write summary informatin into sheet
def write_summary(project, writer_summary):
    global summary_df
    ## Input data in table 'Artifact Finding Types and Count'
    add_row = 1
    size1 = 0
    size2 = 0
    ## Cut the length if the length > 31
    project = project if len(project) <= 31 else project[:30]
    
    for index,find_type in enumerate(find_types_df):
        key = find_type.columns.values[0]
        start = add_row     
        if index % 2 == 0 :
            size1 = len(find_type)
            find_type.to_excel(writer_summary,sheet_name = project,startrow = start,startcol=7,index =False)

            if project_end:
                summary_findtypes_dict[key].to_excel(writer_summary,sheet_name = summary_sh,startrow = start,startcol=7,index =False)
        else:
            size2 = len(find_type)
            find_type.to_excel(writer_summary,sheet_name = project,startrow = start,startcol=10,index=False)
            if project_end:
                summary_findtypes_dict[key].to_excel(writer_summary,sheet_name = summary_sh,startrow = start,startcol=10,index=False)
            
            add_row += max(size1, size2) + 3
        
    ## Insert table name
    worksheet = writer_summary.sheets[project]
    worksheet.write(0, 0, metric_tb)
    worksheet.write(0, 7, type_tb)
    if project_end:
        summary_sheet = writer_summary.sheets[summary_sh]
        summary_sheet.write(0, 0, metric_tb)
        summary_sheet.write(0, 7, type_tb)
        set_format(summary_sheet, writer_summary)

    ## Input data in table 'Findings Metrics'
    metric_df = pd.DataFrame(find_metric_list, columns = metric_head)
    total_row = []
    metric_size = len(metric_df)
    metric_df.to_excel(writer_summary, sheet_name = project,startrow = 1,startcol=0,index=False)
    ## Sum table 'Findings Metrics'
    left_index = metric_df[metric_head[0]]
    metric_df = metric_df.iloc[0:,1:]
    
    if summary_df.empty:
        summary_df = metric_df
    else:
        summary_df = summary_df.add(metric_df,fill_value=0)
    if project_end:
        left_index.to_excel(writer_summary, sheet_name = summary_sh,startrow = 1,startcol=0,index=False)
        summary_df.to_excel(writer_summary, sheet_name = summary_sh,startrow = 1,startcol=1,index=False)

    write_projects(project, writer_summary, metric_df)

    ## Input data in left mid-table
    # Project
    left_list = [['# of Test Cases', count_testcase], ['# of Test Results', count_testresult]]
    # Summary
    left_sum_list = [['# of Test Cases', sum_testcase], ['# of Test Results', sum_tr]]
    
    total_df = pd.DataFrame(left_list)
    total_df.to_excel(writer_summary, sheet_name = project,startrow = 12,startcol=0,index=False,header=False)
    ## Sum left table
    left_df = pd.DataFrame(left_sum_list)
    left_df.to_excel(writer_summary, sheet_name=summary_sh,startrow=12,startcol=0,index=False,header=False)

    set_format(worksheet, writer_summary)


## Write projects data into All projects sheet
def write_projects(project, writer_summary, summary_df):
    global projects_list
    project_list = list(summary_df['Total # of QA'])
    sum_project = sum(project_list)
    project_list.insert(0,project)
    project_list.insert(len(project_list),sum_project)
    projects_list.append(project_list)
    
    if project_end:
        projects_df = pd.DataFrame(projects_list)
        headers = ['Project Reviewed', 'Test Plan', 'Test Cases', 'RTM', 'Test Evidences','Test Summary Report', 'SUM']
        projects_df.columns = headers
        sum_list = []
        for item in headers[1:]:
            count_col = projects_df[item].sum()
            sum_list = sum_list + [count_col]
            
        projects_df.loc[-1] = ['SUM'] + sum_list
        projects_df.index = projects_df.index + 1
        
        projects_df.to_excel(writer_summary, sheet_name=project_sh,startrow=2,startcol=1,index=False)


## Write summary data into category sheet
def write_category(writer_summary):
    print('@@@@@@@@@@@@@@@@@@@@@@@@@@@category')
    ## All category page setting
    header_category = ['Deliverable', 'Observation', 'Total']
    summary_df = pd.DataFrame()
    
    start = True
    for keyvalue in summary_findtypes_dict.items():

        key, value = keyvalue[0], keyvalue[1]
        
        key = key.replace(' Finding Type', '')
        key = 'Test Evidences' if key == 'Test Results' else key
        
        deliveriable = [key]*len(value)
        value.insert(0,header_category[0],deliveriable)
        value.columns = header_category

        # Filter the tracker observations
        obs_col = 'Observation'
        value = value.loc[~value[obs_col].str.lower().isin(tracking_str_list)]

        # Ger rid of Grand total
        value = value[value.Observation != 'Grand Total']

        # Re calculate Grand total
        grand_total = value['Total'].sum() * 1

        value.loc[len(value)] = [key, 'Grand Total', grand_total]
        
        if start:
            summary_df = value
            start = False
        else:
            summary_df = summary_df.append(value)

    
        
    summary_df.to_excel(writer_summary, sheet_name=category_sh,startrow=start,startcol=0,index=False)

    ## Order the summary table
    summary_order_df = summary_df[summary_df[header_category[1]] != 'Grand Total']
    summary_order_df = summary_order_df.sort_values(header_category[2], ascending=False)
    summary_order_df.to_excel(writer_summary, sheet_name=category_sh,startrow=start,startcol=9,index=False)

    ## mid-table
    mid_df = summary_df[summary_df[header_category[1]] == 'Grand Total']
    mid_df = mid_df.sort_values(header_category[2], ascending=True)
    mid_df = mid_df.reset_index(drop=True)
    total = mid_df[header_category[2]].sum()
    mid_df['percent'] = (mid_df[header_category[2]] / total)
    percent = mid_df['percent'].sum()

    mid_df.loc[-1] = [np.nan,np.nan,total,percent]
    mid_df['percent'] = pd.Series(["{0:.2f}%".format(val * 100) for val in mid_df['percent'] ], index = mid_df.index)
    mid_df.index = mid_df.index + 1

    mid_df.to_excel(writer_summary, sheet_name=category_sh,startrow=1,startcol=4,index=False,header=False)


def write_review(projects,result_list):
    print(projects, result_list)
    writer = pd.ExcelWriter(review_file,engine='xlsxwriter')
    workbook = writer.book

    tb_merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'gray'})

    ## header
    header_df = pd.DataFrame([review_header])
    header_df.to_excel(writer, sheet_name = review_sh, startrow=0, startcol=0, index=False, header=False)
    header_cols = []
    header_rows = []
    header = []
    for i in review_cols:
        header_cols += [i, '', '']
        header_rows += review_sub
    header.append(header_cols)
    header.append(header_rows)
    header_df2 = pd.DataFrame(header)
    header_df2.to_excel(writer,sheet_name=review_sh,startrow=0, startcol=12,index=False,header=False)

    worksheet = writer.sheets[review_sh]
    index = 3
    for name in projects:
        worksheet.write('D'+str(index),name)
        index += 1
        

    ## Result
    row_df = pd.DataFrame(result_list)
    row_df.to_excel(writer,sheet_name = review_sh,startrow = 2,startcol=10,index =False,header = False)
        
    writer.save()

def set_vars(input_start_date,input_end_date,input_copy_path,input_project_path,input_dir_path,sheet_name_field_value,input_summary_file,input_review_file,input_review_sh,input_service_file):
    global dir_path, start_date, end_date, copy_path,service_file,sheets
    global summary_file, review_file, review_sh, review_cols, project_path

    start_date = datetime.datetime.strptime(input_start_date,'%m/%d/%Y')
    end_date = datetime.datetime.strptime(input_end_date,'%m/%d/%Y')
    print(start_date,end_date)
    copy_path = input_copy_path  
    project_path = input_project_path

    dir_path = input_dir_path

    summary_file = input_summary_file
    review_file = input_review_file
    review_sh = input_review_sh
    service_file = input_service_file
    sheets = sheet_name_field_value.split(",")

def read(project_dict_qa):
    global project_dict,finding_type_df, find_types_df, find_metrics_df, find_metric_list
    global processed_data_dict, count_testresult, count_testcase
    global summary_findtypes_dict, review_list, project_end
    global total_opened, total_closed, total_remain
    project_dict = dict(project_dict_qa)
    read_files_software(project_dict)
    process()

def process():
    global project_dict,finding_type_df, find_types_df, find_metrics_df, find_metric_list
    global processed_data_dict, count_testresult, count_testcase
    global summary_findtypes_dict, review_list, project_end
    global total_opened, total_closed, total_remain

    print("-----------------------Process" , project_dict)
    ## Create the result excel and worksheet
    writer_summary = pd.ExcelWriter(summary_file,engine='xlsxwriter')

    # Define order of sheets
    empty_df = pd.DataFrame()
    empty_df.to_excel(writer_summary,sheet_name = summary_sh)
    empty_df.to_excel(writer_summary,sheet_name = project_sh)
    empty_df.to_excel(writer_summary,sheet_name = category_sh)

    project_num = len(project_dict)

    ## initiate list for each sheet
    for name in types_dict:
        summary_findtypes_dict[name + ' Finding Type'] = pd.DataFrame()

    result_lists = []
    
    for project in project_dict:
        project_num -= 1
        if project_num == 0:
            project_end = True

        ## Reset variables
        processed_data_dict = {}
        finding_type_df = []
        find_types_df = []
        find_metrics_df = []
        find_metric_list = []
        count_testcase = 0
        count_testresult = 0

        # Tracking observation variables

        
        print('**'+project)
        print("-----Project----\t", project)
        get_valid_type(project)
        write_processed_data(project)
        write_summary(project, writer_summary)
        ## feedback review report
        result_list = [count_testcase, count_testresult] + review_list + [total_opened, total_closed, total_remain]
        result_lists.append(result_list)
        review_list = []
        total_opened = 0
        total_closed = 0
        total_remain = 0
        opened_obs = []
        print('*************************************')
        
    write_category(writer_summary)
    
    writer_summary.save()
    writer_summary.close()
    
    print("----------NEW",list(project_dict.keys()))
    print("---TWo", result_lists)
    write_review(list(project_dict.keys()),result_lists)

def aakash_script_charts_pie():
    output_dir_path = dir_path

    file = 'qa_review_summary.xlsx' # file name
    full_input_path = os.path.join(dir_path,file) 
    top = 5 # 2
    output_file = 'Findings_Report_' + file
    full_output_path = os.path.join(output_dir_path,output_file) 
    write_sheet = 'Sheet1'


    # Output Excel file for charts.
    startcol = 'C'
    strow = 10

    # Color settings for Plan, Cases & Results/Evidences.
    colors = ['#40BCD8','#18AD91','#1C77C3']
    # colors = {'Test Plan': '#40BCD8', 'Test Cases' : '#18AD91' , 'Test Results' : '#1C77C3'}  ## Test Plan, Test Cases, Test Evidences


    # ### Number of sheets to be converted into charts.

    # In[5]:


    xl = pd.ExcelFile(full_input_path)
    summary = xl.sheet_names[0]        #'Summary' #0
    all_projects = xl.sheet_names[1]   #'All projects' #1
    all_categories = xl.sheet_names[2] #'All categories' #2
    projects = xl.sheet_names[3:]
    print(projects)


    # In[6]:


    for i in projects:
        print(i)
        df = xl.parse(i,header=1, usecols="A:E", nrows=6) #index_col=0
        df = df.transpose()
        df.columns = df.iloc[0]
        
        #Removing unwanted columns
        df = df.drop(df.columns[2],1) #'RTM'
        df = df.drop(df.columns[4],1) #'Test Summary Report'
        df = df.drop(df.columns[3],1) #'Requirements'
        #TODO:remove duplication
        
        print(df)
        print("\n")


    # ## Data Manipulation

    # ### Summary Sheet

    # In[7]:


    sheet_name = summary
    df = xl.parse(sheet_name,header=1, usecols="A:E", nrows=6) #index_col=0
    print(df)


    # In[8]:


    df = df.transpose()
    df.columns = df.iloc[0]
    print(df)


    # In[9]:


    #Removing unwanted columns
    df = df.drop(df.columns[2],1) #'RTM'
    df = df.drop(df.columns[4],1) #'Test Summary Report'
    df = df.drop(df.columns[3],1) #'Requirements'
    print(df)
    #TODO:remove duplication


    # ### Output refined data into Excel and format it.

    # In[11]:

    writer = pd.ExcelWriter(full_output_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=write_sheet, startrow=15, startcol=8)
    workbook = writer.book
    worksheet = writer.sheets[write_sheet]

    format = workbook.add_format()
    format.set_align('center')
    format.set_align('vcenter')
    worksheet.set_column('I:L',20, format)


    # ### Chart 1: Pie Chart for breakdown of test cases, plans and evidences.

    # In[17]:


    # Create a pie chart object.
    chart = workbook.add_chart({'type': 'pie'})

    categories = '='+ write_sheet +'!J17:L17'
    values = '='+ write_sheet +'!J18:L18'

    chart.add_series({
        'name' : 'pie_series',
        'categories': categories,
        'values':     values,
        'points': [
            {'fill': {'color': colors[0]}},
            {'fill': {'color': colors[1]}},
            {'fill': {'color': colors[2]}}
        ],
        'data_labels': {'value': True},
    })

    chart.set_size({'width': 620, 'height': 456})
    chart.set_title ({
        'name': 'Total number of findings per deliverable type'
    })

    # Set an Excel chart style. Colors with white outline and shadow.
    chart.set_style(10)

    # Insert the chart into the worksheet.
    format_chart = workbook.add_format()
    worksheet.insert_chart('A23', chart,{'y_offset': -150}) #{'x_offset': 5, 'y_offset': 5}
    worksheet.set_column('A:E',15, format_chart)


    # In[21]:


    #Create_Bar_Chart_For_Top_Findings
    sheet_name = all_categories
    df = xl.parse(sheet_name,header=0, usecols="J:L", nrows=top)

    df['Combo'] = df.Deliverable.str.cat(" - " + df.Observation)
    df = df.sort_values(by=['Total'],ascending=True)
    df.to_excel(writer, sheet_name=write_sheet, startrow=0, startcol=0, index=False,header=False)
    workbook = writer.book
    worksheet = writer.sheets[write_sheet]

    fmt = writer.book.add_format({'font_color': 'black'})
    worksheet.set_row(1,14,fmt)
    worksheet.set_column('A:H', 14, fmt)

    print(df)


    # ### Chart 2: Bar chart (top) for top findings in a year.

    # In[24]:


    chart2 = workbook.add_chart({'type': 'bar'})

    # Configure the first series.

    points = []
    for index, row in df.iterrows():
       clr = {}
       fill = {}
       if(row['Combo'].startswith('Test Evidences')):
        clr = {'color' : colors[2] }
       elif(row['Combo'].startswith('Test Cases')):
        clr = {'color' : colors[1] }
       elif(row['Combo'].startswith('Test Plan')):
        clr = {'color' : colors[0] }

       fill = {'fill' : clr}
       points.append(fill)

    chart2.add_series({
        'name':       'top_series',
        'categories': [write_sheet, 0,3,0+top ,3],
        'values':     [write_sheet, 0,2,(0+top),2],
        'data_labels': {'value': True},
        'points':points
    })


    # Add a chart title and some axis labels.
    chart2.set_size({'width': 720, 'height': 256})
    chart2.set_legend({'none': True})
    chart2.set_title ({'name': 'Top ' + str(top) + ' Deliverables'})
    chart2.set_x_axis({'major_gridlines': {'visible': False},'visible': False})#'name': 'Total Findings in category',
    chart2.set_y_axis({'major_gridlines': {'visible': False}, 'num_font':  {'bold': True}}) #'name': 'Categories',

    # Set an Excel chart style.
    chart2.set_style(10)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('F1', chart2)


    # In[25]:


    writer.save()
    workbook.close()
    writer.close()

def set_quarter_vars(input_source_path,input_dir_path,input_number_sheets,input_top_n):
    global dir_path_yearly, source_path_yearly, number_sheets, top_n

    source_path_yearly = input_source_path
    dir_path_yearly = input_dir_path
    number_sheets = input_number_sheets
    top_n = input_top_n

def aakash_script_charts_column():
    sheets = number_sheets
    top = top_n # 2

    # Output Excel file for charts.
    output_file = os.path.join(os.path.dirname(dir_path_yearly), 'bar_report_' + os.path.basename(source_path_yearly))
    write_sheet = 'Sheet1'
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    startcol = 'C'
    strow = 10


    # Color settings for Plan, Cases & Results/Evidences.
    colors = {'Test Plan':'#40BCD8', 'Test Cases' : '#18AD91' , 'Test Results' : '#1C77C3'}  ## Test Plan, Test Cases, Test Evidences

    # ### Number of sheets to be converted into charts.
    xl = pd.ExcelFile(source_path_yearly)
    years = xl.sheet_names[0:sheets]
    print(years)


    # ### Data Manipulation
    # 
    # ##### Example sheet looks like
    # 
    # ![image.png](attachment:image.png)

    # In[6]:


    #Helper functions.

    def getQuarterDataFrame(year):
        df = xl.parse(year,usecols="A:B")
        df = df.dropna()
        quarter_list = df.index.tolist()    
        quarter_dict = {}
        
        for idx, quarter in enumerate(quarter_list):
            Q = 'Q'
            Q = Q + str(idx+1)
            quarter_dict[Q] = [quarter_list[idx]]

        df = pd.DataFrame(data=quarter_dict);
        
        quarter_list = df.index.tolist()
        idx = quarter_list.index(0)
        quarter_list[idx] = str(year)
        df.index = quarter_list
        
        return df

    def getFindingsDataFrame(year):
        df2 = xl.parse(year, header=0, usecols="B:E", skipfooter=1)
        total_findings[year] = df2[['Total']].sum(axis=0)
        df2['obs_total'] = df2.groupby(['Observation'])['Total'].transform('sum')
        df2 = df2.drop_duplicates(['Observation'])
        df2 = df2.drop(['Quarter','Total'],axis=1)
        df2 = df2.sort_values(by=['obs_total'], ascending=False)
        df2 = df2.iloc[0:top]
        return df2

    def checkEvidencesResults(deliverable_list):
            new_items = ['Test Results' if x == 'Test Results' or x =='Test Evidences' else x for x in deliverable_list]
            return new_items


    # In[7]:


    df = []
    top_finding = {}
    top_finding_deliverable = {}
    top_finding_observation = {}
    total_findings = {}

    for year in years:
        
        #Dictionary entry for each year.
        top_finding[year] = [];
        
        #Parse yearly to get totals of each quarter.
        #              Q1     Q2     Q3     Q4
        #    FY17  2154.0   72.0  504.0  827.0
        #    FY18   827.0  725.0  715.0  573.0
        #    FY19   287.0    NaN    NaN    NaN 

        df1 = getQuarterDataFrame(year)
        df.append(df1)
        
        #Parse yearly to get top findings in the year.
        #       FY17  FY18  FY19
        #    0  1320  1047   118
        #    1   841   601   106 
        
        df2 = getFindingsDataFrame(year)
        
        top_finding[year] = df2['obs_total'].tolist();
        new_items = checkEvidencesResults(df2['Deliverable'].tolist())
        
        #Extra Information.
        top_finding_deliverable[year] = new_items;
        top_finding_observation[year] = df2['Observation'].tolist();

        
    df1 = pd.concat(df, axis=0,sort=True)
    print(df1, "\n")

    df2 = pd.DataFrame(data=top_finding)
    print(df2 , "\n")

    df3 = pd.DataFrame(data=top_finding_deliverable)
    print(df3 , "\n")

    df4 = pd.DataFrame(data=top_finding_observation)
    print(df4 , "\n")

    df5 = pd.DataFrame(data=total_findings)
    print(df5 , "\n")

    # Important
    df5 = df5.T


    # ### Output refined data into Excel and format it.

    # In[8]:


    df1.to_excel(writer, sheet_name=write_sheet, startrow=0, startcol=0)
    df2.to_excel(writer, sheet_name=write_sheet, startrow=0, startcol=6)
    # df3.to_excel(writer, sheet_name=write_sheet, startrow=20, startcol=col_start)    

    workbook = writer.book
    worksheet = writer.sheets[write_sheet]

    format = workbook.add_format()
    format.set_align('center')
    format.set_align('vcenter')
    worksheet.set_column('A:Z',10, format)


    # ### Chart 1: Stacked chart (left) for Quarter/Year comparison.

    # In[9]:


    chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

    # Configure the series of the chart from the dataframe data.
    for col_num in range(1, 5):
        chart.add_series({
            'name': [write_sheet, 0, col_num, 0, col_num],
            'categories': [write_sheet, 0, col_num, 0, col_num],
            'values':     [write_sheet, 1, col_num, len(years), col_num],
            'gap':        10,
            'data_labels': {'value': True, 'category': True},
            'fill':       {'color': brews['Pastel1'][col_num]},
        })

    # Configure the chart axes.
    chart.set_title ({'name': 'Total Findings Testing Artifacts'})
    chart.set_x_axis({'major_gridlines': {'visible': False},'visible': False})#'name': 'Total Findings in category',
    chart.set_y_axis({'major_gridlines': {'visible': False},'visible': False}) #'name': 'Categories',

    # Insert the chart into the worksheet.
    worksheet.insert_chart('A15', chart)


    # ### Chart 2: Column chart (middle) for Top findings in a year.

    # In[10]:


    n = len(years) - 1
    col_start = 6
    col_num_start = col_start + 1
    row_num = 15
    mx = max(df2.max())

    for i in range(1,top+1):
        chart = workbook.add_chart({'type': 'column'})

        categories = [write_sheet, 0, col_num_start, 0, col_num_start+n]
        values = [write_sheet, i, col_num_start, i, col_num_start+n]
        points = []
        
        for idx,series in df3.iterrows():
            if(idx == i-1):
                for x in series:
                    points.append({'fill' : {'color' : colors[x] }}) 

        # Configure the series of the chart from the dataframe data.
        chart.add_series({
            'name': 'Test',
            'categories': categories,
            'values':     values,
            'gap':        10,
            'data_labels': {'value': True, 'category': True},
            'points':points
        })

        # Configure the chart axes.
        chart.set_size({'width': 256, 'height': 256})
        chart.set_title ({'name': 'Top ' + str(i) + ' finding '})
        chart.set_legend({'none': True})
        chart.set_x_axis({'major_gridlines': {'visible': False},'visible': False})#'name': 'Total Findings in category',
        chart.set_y_axis({'major_gridlines': {'visible': False},'visible': False,'min': 0, 'max': mx}) #'name': 'Categories',
        chart.set_style(25)

        # Insert the chart into the worksheet.
        worksheet.insert_chart(row_num,7, chart)
        row_num = row_num + 15

        


    # ### Table 1: Simple Table (right) for Total findings.

    # In[11]:


    caption = 'Default table with no data.'

    # Set the columns widths.
    worksheet.set_column('B:G', 12)

    # Write the caption.
    worksheet.write('B1', caption)

    # Add a table to the worksheet.
    worksheet.add_table(10,11,10+n+1,12)

    df5.to_excel(writer, sheet_name=write_sheet, startrow=10, startcol=11)    
    workbook = writer.book
    worksheet = writer.sheets[write_sheet]


    # In[12]:


    writer.save()
    writer.close()



## Format the sheets of QA Metric summary and projects
def set_format(worksheet, writer_summary):
    # Create a format for each table
    workbook = writer_summary.book
    count_format = workbook.add_format({'num_format':'format10'})
    tb_merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'fg_color': 'gray'})
    tb_head_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'fg_color': '#8db4e2'})
    

    ## Merge cells
    merge_range1 = 'A1:'+str(chr(ord('A')+(len(metric_head))-1))+'1'
    merge_range2 = 'H1:'+str(chr(ord('H')+(len(metric_head )+5)-1))+'1'
    worksheet.merge_range(merge_range1, metric_tb, tb_merge_format)
    worksheet.merge_range(merge_range2, type_tb, tb_merge_format)
