#Importing required Python modules
import pandas as pd
import xlrd 
import datetime
import matplotlib.pyplot as plt
import numpy as np
import argparse
from collections import defaultdict
from matplotlib.backends.backend_pdf import PdfPages
import os

#Fetching Command Line arguments
def fetch_cmdline_options():
    parser = argparse.ArgumentParser(description='Env Availability report')
    parser.add_argument('--start_date', dest='start_date', required = True, 
                        help='Enter the start date in format mm/dd/yyyy')
    parser.add_argument('--end_date', dest='end_date', required = True,
                        help='Enter the end date in format mm/dd/yyyy')
    parser.add_argument('--holiday_file', dest='holiday_file', default='holiday.txt',
                        help='provide file containing holidays in format mm/dd/yyyy')
    parser.add_argument('--input_file', dest='input_file', default= 'task_env_report.xlsx',
                        help='provide input excel file containing env availability data')
    args = parser.parse_args()
    start_date, end_date, input_file, holiday_file = args.start_date, args.end_date, args.input_file, args.holiday_file
    
    #Converting start and end date to datetime objects
    start_date = datetime.datetime.strptime(start_date, '%m/%d/%Y')
    end_date = datetime.datetime.strptime(end_date, '%m/%d/%Y')
    
    #Initial input checks
    if start_date > end_date:
        print("Start Date is greater than end date, therefore, exiting the program")
        exit(0)
    if not os.path.isfile(input_file):
        print("Input File :{0} doesn't exist".format(input_file))
        exit(0)
    if not os.path.isfile(holiday_file):
        print("Holiday File :{0} doesn't exist".format(holiday_file))
        exit(0)
    
    return start_date, end_date, input_file, holiday_file
 
#Fetching data from the excel
def fetch_input_data(input_file):
    df_sheet_data = pd.read_excel(input_file, sheet_name='env_data')  
    df_sheet_data['start_time'] = [d.time() for d in df_sheet_data['start_date']]
    df_sheet_data['end_time'] = [d.time() for d in df_sheet_data['end_date']]
   
    return df_sheet_data

#Split the downtimes for different dates in case the time exceeds 24 hours
def split_date(Environment,Category, Planned, start_date, end_date, summary, start_time, end_time):
    #Same day case                                                                                 
    if start_date == end_date:                                                                 
        return [(Environment, Category, Planned, start_date, end_date,  summary, start_time, end_time)]                                                                      

    stop_split = (datetime.datetime.combine(datetime.date(1,1,1),start_time.replace(hour=0, minute=0, second=0)) + datetime.timedelta(days=1)).time()
    return [(Environment, Category, Planned, start_date,end_date, summary, start_time, stop_split)] + split_date(Environment,Category, Planned, end_date, end_date,  summary, stop_split, end_time)

#Calculating the dates for which report needs to be generated 
def calculate_reporting_date_range(start_date, end_date, holiday_file):
    report_dates_24hr = pd.date_range(start_date, end_date)
    report_dates_workday = pd.bdate_range(start_date, end_date)
    
    #Excluding the holidays for workday plot
    holiday_file = open(holiday_file, "r")
    dates_to_be_excluded = holiday_file.read()
    dates_to_be_excluded = dates_to_be_excluded.strip().split("\n")

    for dt in dates_to_be_excluded:
        dt = datetime.datetime.strptime(dt, '%m/%d/%Y')
        if dt >= start_date and dt <= end_date:
            report_dates_workday = report_dates_workday.drop(dt)
    holiday_file.close()
   
    return report_dates_24hr, report_dates_workday
    
def filter_input_data(df_sheet_data, start_date, end_date, holiday_file):
    splitted_dates = [
        elt for _, row in df_sheet_data.iterrows() for elt in split_date(row["Environment"],row["Category"], row["Planned"], row["start_date"].date(), row["end_date"].date(), row["summary"], row["start_time"], row["end_time"])
    ]      

    #Here, we have data for each individual date now
    filtered_dataframe = pd.DataFrame(splitted_dates, columns=list(df_sheet_data.columns))

    #Filter out the rows for which we need the report for
    report_dates_24hr, report_dates_workday = calculate_reporting_date_range(start_date, end_date, holiday_file)
    
    filtered_dataframe_24hr = filtered_dataframe.loc[(filtered_dataframe['start_date'].isin(report_dates_24hr.date))]
    filtered_dataframe_workday = filtered_dataframe.loc[(filtered_dataframe['start_date'].isin(report_dates_workday.date))]   
   
    
    days_24hr = len(filtered_dataframe_24hr['start_date'].unique())
    days_workday = len(filtered_dataframe_workday['start_date'].unique())

    return filtered_dataframe_24hr, filtered_dataframe_workday, days_24hr, days_workday

#Calculating the workday downtime
def calculate_workday_downtime(filtered_dataframe):
    from_time = datetime.datetime.strptime('08:00', '%H:%M').time()
    to_time = datetime.datetime.strptime('17:00', '%H:%M').time()
    
    #breaking the time ranges based on workday timings and considering only those that falls between the workday timing.
    for i, row in filtered_dataframe.iterrows():
        from_datetime = datetime.datetime.combine(row['start_date'], from_time)
        to_datetime = datetime.datetime.combine(row['start_date'], to_time)
        stime = row['start_time']
        etime = row['end_time']
        sdatetime = datetime.datetime.combine(row['start_date'], row['start_time'])
        edatetime = datetime.datetime.combine(row['start_date'], row['end_time'])
        
        if row['start_date'] == row['end_date']:
            if stime < from_time and etime < from_time:
                filtered_dataframe.at[i, 'downtime_workday'] = 0
        
            elif stime > to_time and etime > to_time:
                filtered_dataframe.at[i, 'downtime_workday'] = 0 
        
            elif stime < from_time and etime > to_time:
                diff = (to_datetime - from_datetime)/datetime.timedelta(minutes=1)
                filtered_dataframe.at[i, 'downtime_workday'] = diff    
        
            elif stime >= from_time and etime > to_time:
                diff = (to_datetime - sdatetime) /datetime.timedelta(minutes=1)
                filtered_dataframe.at[i, 'downtime_workday'] = diff
        
            elif stime < from_time and etime <= to_time:
                diff = (edatetime - from_datetime)/datetime.timedelta(minutes=1)
                filtered_dataframe.at[i, 'downtime_workday'] = diff
        
            elif stime >= from_time and etime <= to_time:
                diff = (edatetime - sdatetime)/datetime.timedelta(minutes=1)
                filtered_dataframe.at[i, 'downtime_workday'] = diff
        
        else:
            if stime > to_time:
                filtered_dataframe.at[i, 'downtime_workday'] = 0
            elif stime >= from_time and stime <=to_time:
                diff = (to_datetime - sdatetime)/datetime.timedelta(minutes=1)
                filtered_dataframe.at[i, 'downtime_workday'] = diff
            elif stime < from_time :
                diff = (to_datetime - from_datetime)/datetime.timedelta(minutes=1)
                filtered_dataframe.at[i, 'downtime_workday'] = diff

    return filtered_dataframe

#Calculation of downtime
def calculate_downtime(filtered_dataframe_24hr, filtered_dataframe_workday):
   
    #Calculate data for 24 hour
    start_time = pd.to_datetime(filtered_dataframe_24hr['start_time'].astype(str))
    end_time = pd.to_datetime(filtered_dataframe_24hr['end_time'].astype(str))
    day_diff = filtered_dataframe_24hr['end_date'] - filtered_dataframe_24hr['start_date']
    downtime = (end_time - start_time) + day_diff
    filtered_dataframe_24hr['downtime'] = downtime / datetime.timedelta(minutes=1)
    
    filtered_dataframe_24hr = filtered_dataframe_24hr[filtered_dataframe_24hr['downtime'] != 0]

    #Calculate downtime data for workday plots
    filtered_dataframe_workday = calculate_workday_downtime(filtered_dataframe_workday)
    print(filtered_dataframe_workday)
    return filtered_dataframe_24hr, filtered_dataframe_workday

#calculation of uptime and dowtime statisitics for plotting
def calculate_statistics(filtered_dataframe_24hr, filtered_dataframe_workday, environments, days_24hr, days_workday):
    #dictionary to collect statistics
    stats_dict = defaultdict(list)
    stats_dict_workday = defaultdict(list)
       

    #Calculate the statistics for each env
    for env in environments:
        if env not in stats_dict: 
            pct_planned_downtime = (filtered_dataframe_24hr[filtered_dataframe_24hr['Environment'] == env][filtered_dataframe_24hr['Planned'] == 'Yes']['downtime'].sum()/(1440*days_24hr)) *100 
            pct_unplanned_downtime = (filtered_dataframe_24hr[filtered_dataframe_24hr['Environment'] == env][filtered_dataframe_24hr['Planned'] == 'No']['downtime'].sum()/(1440*days_24hr))* 100 
            pct_uptime = 100 - pct_planned_downtime - pct_unplanned_downtime
            summary = "\n".join(list(filtered_dataframe_24hr[filtered_dataframe_24hr['Environment'] == env]['summary']))
            stats_dict[env] = [pct_uptime, pct_planned_downtime, pct_unplanned_downtime, summary]
        else:
            print("{0} Environment already exists".format(env))
        
        if env not in stats_dict_workday:
            pct_planned_downtime = (filtered_dataframe_workday[filtered_dataframe_workday['Environment'] == env][filtered_dataframe_workday['Planned'] == 'Yes']['downtime_workday'].sum()/(1440*days_workday)) *100 
            pct_unplanned_downtime = (filtered_dataframe_workday[filtered_dataframe_workday['Environment'] == env][filtered_dataframe_workday['Planned'] == 'No']['downtime_workday'].sum()/(1440*days_workday))* 100 
            pct_uptime = 100 - pct_planned_downtime - pct_unplanned_downtime
            summary = "\n".join(list(filtered_dataframe_workday[filtered_dataframe_workday['Environment'] == env]['summary']))
            stats_dict_workday[env] = [pct_uptime, pct_planned_downtime, pct_unplanned_downtime, summary]              

    return stats_dict, stats_dict_workday

#Function to generate plots and dump it to a PDF 
def generate_plots(labels, colors, stats_dict,stats_dict_workday, environments, start_date, end_date):

    with PdfPages(r'Env_Availability_report.pdf') as export_pdf:
        rows = 2
        cols = len(environments)
        fig, ax = plt.subplots(rows, cols)
        title = "Workday"
        chart_data = stats_dict_workday

        for row in range(rows):
            for col, env in enumerate(environments):
                ax[row, col].pie(chart_data[env][:3], colors=colors, shadow = False, startangle=90, autopct=lambda p: '{:.2f}%'.format(round(p)) if p > 0 else '')
                ax[row, col].axis('equal')
                ax[row, col].text(0, 0, '{0} {1}'.format(env, title), va = 'center', ha = 'center', fontsize=10, fontweight="bold")
                
                fig = plt.gcf()
                fig.set_size_inches(15, 10)
                circle = plt.Circle(xy=(0,0), radius=0.80, facecolor='white')
                rec = plt.Rectangle((-1.1, -1.1),  2.2, 2.19, fill=False, lw=3, zorder=100)
                ax[row, col].add_patch(rec)
                ax[row, col].add_patch(circle)
            title = "24 Hr"
            chart_data = stats_dict
        
        fig.suptitle("Dev and Clone Environment Availability (Date Range: {0} to {1})".format(start_date.date(), end_date.date()),fontsize=14, fontweight="bold")
        fig.legend(labels=[f'{x}' for x in labels], bbox_to_anchor=(0.66, 0.95), prop={'size': 10}, ncol = 3)
        fig.tight_layout(h_pad=2, w_pad=-5) 
        fig.subplots_adjust(top=0.93)
        export_pdf.savefig()
        plt.show()
        plt.close()

 
if __name__ == "__main__":
    print("Start of program\n")
    environments = ['Development', 'Clone Pre Prod','Clone UAT']
    labels= ['Uptime', 'Planned Downtime', 'Unplanned Downtime']
    colors=['green', 'orange', 'red']
    
    start_date, end_date, input_file, holiday_file = fetch_cmdline_options()
    df_sheet_data = fetch_input_data(input_file)
    filtered_dataframe_24hr, filtered_dataframe_workday, days_24hr, days_workday = filter_input_data(df_sheet_data, start_date, end_date, holiday_file)
    filtered_dataframe_24hr, filtered_dataframe_workday = calculate_downtime(filtered_dataframe_24hr, filtered_dataframe_workday)

    stats_dict, stats_dict_workday = calculate_statistics(filtered_dataframe_24hr, filtered_dataframe_workday, environments, days_24hr, days_workday)
    generate_plots(labels, colors, stats_dict,stats_dict_workday, environments, start_date, end_date)

    print("\nEnd of Program")