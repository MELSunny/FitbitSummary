import os
import pandas as pd
import time
from datetime import datetime
import xlsxwriter
import math
PATH=os.path.join(os.getcwd(),"Data0816") #Default


def process_datetime(range_date_start,range_date_end,date_default):
    while True:
        user_input= input("(format: DD/MM/YYYY, default: %s):" % date_default.strftime('%d/%m/%Y'))
        if user_input != '':
            try:
                input_date = datetime.strptime(user_input, '%d/%m/%Y').date()
            except ValueError:
                print("Incorrect date format, should be DD/MM/YYYY!")
            else:
                if input_date <range_date_start or input_date>range_date_end:
                    print("Incorrect date range, should be after %s and before %s" % (range_date_start.strftime('%d/%m/%Y'),range_date_end.strftime('%d/%m/%Y')))
                else:
                    return input_date
        else:
            return date_default


print('Current data path:%s'% PATH)
user_input=input('If it is right, please press ENTER key, otherwise input the path:')
if user_input!='':
    while True:
        if (os.path.exists(user_input)):
            PATH=user_input
            break
        else:
            user_input = input('Invalid path! Please try again:')
csvfiles = [f for f in os.listdir(PATH) if os.path.isfile(os.path.join(PATH, f)) and f.endswith('.fitbit.csv')]
workbook = xlsxwriter.Workbook('result.xlsx')
worksheet = workbook.add_worksheet()
merge_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter'})

worksheet.merge_range('A1:A2', 'Patient ID',merge_format)
worksheet.merge_range('B1:B2', 'Start date',merge_format)

worksheet.merge_range('C1:G1', 'Before operation per day',merge_format)
worksheet.merge_range('H1:H2', 'Operation date',merge_format)
worksheet.merge_range('I1:M1', 'After operation per day',merge_format)
worksheet.merge_range('N1:N2', 'End date',merge_format)
worksheet.write('C2', 'Avg Steps',merge_format)
worksheet.write('D2', 'Avg Distance',merge_format)
worksheet.write('E2', 'Avg Elevation',merge_format)
worksheet.write('F2', 'Avg CaloriesOut',merge_format)
worksheet.write('G2', 'Avail of days',merge_format)

worksheet.write('I2', 'Avg Steps',merge_format)
worksheet.write('J2', 'Avg Distance',merge_format)
worksheet.write('K2', 'Avg Elevation',merge_format)
worksheet.write('L2', 'Avg CaloriesOut',merge_format)
worksheet.write('M2', 'Avail of days',merge_format)


row = 2
for csvfile in csvfiles:
    # csvfile=csvfiles[7]
    print('Opening %s file'% csvfile)
    df = pd.read_csv(os.path.join(PATH, csvfile), sep=',', header=0)
    df['date']=pd.to_datetime(df['date'], format='%Y-%m-%d').dt.date
    df_presorted=df.sort_values('steps',ascending=False)
    df_dropped=df_presorted.drop_duplicates(subset='date')
    df_sorted=df_dropped.sort_values('date')
    print('Please input start date')
    start_date=process_datetime(df_sorted.date.iloc[0],df_sorted.date.iloc[-1],df_sorted.date.iloc[0])
    print('Please input end date')
    end_date=process_datetime(start_date, df_sorted.date.iloc[-1], df_sorted.date.iloc[-1])
    print('Please input operation date')
    mid_date=(start_date+(end_date - start_date) / 2)
    operation_date = process_datetime(start_date, end_date, mid_date)
    # print('debug: get date %s    %s    %s'%(start_date,end_date,operation_date,))
    df_deleted=df_sorted[df_sorted["steps"] > 100]

    df_before=df_deleted[df_deleted['date']<operation_date]
    df_before=df_before[df_before['date']>start_date]
    df_after=df_deleted[df_deleted['date']>operation_date]
    df_after = df_after[df_after['date'] < end_date]
    count_before,_=df_before.shape
    avg_steps_before=df_before['steps'].mean()
    avg_distance_before = df_before['distance'].mean()
    avg_elevation_before = df_before['elevation'].mean()
    avg_caloriesOut_before = df_before['caloriesOut'].mean()
    count_after,_=df_after.shape
    avg_steps_after=df_after['steps'].mean()
    avg_distance_after = df_after['distance'].mean()
    avg_elevation_after = df_after['elevation'].mean()
    avg_caloriesOut_after = df_after['caloriesOut'].mean()
    worksheet.write_string(row, 0, csvfile[:-11])
    worksheet.write_datetime(row, 1, start_date)
    if(not math.isnan(avg_steps_before)):
        worksheet.write_number(row, 2, avg_steps_before)
    if (not math.isnan(avg_distance_before)):
        worksheet.write_number(row, 3, avg_distance_before)
    if (not math.isnan(avg_elevation_before)):
        worksheet.write_number(row, 4, avg_elevation_before)
    if (not math.isnan(avg_caloriesOut_before)):
        worksheet.write_number(row, 5, avg_caloriesOut_before)
    worksheet.write_number(row, 6, count_before)
    worksheet.write_datetime(row, 7, mid_date)

    if (not math.isnan(avg_steps_after)):
        worksheet.write_number(row, 8, avg_steps_after)
    if (not math.isnan(avg_distance_after)):
        worksheet.write_number(row, 9, avg_distance_after)
    if (not math.isnan(avg_elevation_after)):
        worksheet.write_number(row, 10, avg_elevation_after)
    if (not math.isnan(avg_caloriesOut_after)):
        worksheet.write_number(row, 11, avg_caloriesOut_after)
    worksheet.write_number(row, 12, count_after)
    worksheet.write_datetime(row, 13, end_date)
    row=row+1

workbook.close()
