#Importing required Python modules
import pandas as pd
import xlrd 
import datetime
import matplotlib.pyplot as plt
import numpy as np
import argparse
from collections import defaultdict
from matplotlib.backends.backend_pdf import PdfPages


#Fetching Command Line arguments
def fetch_cmdline_options():
    parser = argparse.ArgumentParser(description='Env Availability report')
    parser.add_argument('--start_date', dest='start_date', 
                        help='Enter the start date in format mm/dd/yyyy')
    parser.add_argument('--end_date', dest='end_date', 
                        help='Enter the end date in format mm/dd/yyyy')
    parser.add_argument('--holiday_file', dest='holiday_file',
                        help='provide file containing holidays in format mm/dd/yyyy')
    parser.add_argument('--input_file', dest='input_file',# required = True,
                        help='provide input excel file containing env availability data')
    args = parser.parse_args()
    
    start_date, end_date = args.start_date, args.end_date
    
    #Initial input checks
    if (start_date and end_date) and (start_date > end_date):
        print("Start Date is greater than end date, therefore, exiting the program")
        exit(0)

    #Calculating the dates for which report needs to be generated
    report_dates = pd.date_range(start_date, end_date)

    #Excluding the holiday days
    if args.holiday_file:
        holiday_file = open(args.holiday_file, "r")
        dates_to_be_excluded = holiday_file.read()
        dates_to_be_excluded = dates_to_be_excluded.strip().split("\n")
        report_dates = report_dates.drop(dates_to_be_excluded[:])
        holiday_file.close()
    
    return report_dates, args.input_file

#Fetching and filtering data from the excel
def fetch_input_data(report_dates, input_file):
    #input_file = '/content/drive/MyDrive/task_env_report (1).xlsx'
    df_sheet_data = pd.read_excel(input_file, sheet_name='Sheet1')
    df_sheet_data = df_sheet_data[df_sheet_data['start_date'].isin(report_dates)]
    df_sheet_data = df_sheet_data[df_sheet_data['end_date'].isin(report_dates)]
    min_date = min(df_sheet_data['start_date'].dt.date)
    max_date = max(df_sheet_data['start_date'].dt.date)
    no_of_days = (max_date - min_date).days
    return df_sheet_data, str(min_date), str(max_date), no_of_days
    
#Calculation of downtime
def calculate_downtime(df_sheet_data):
    #Calculate data for workday plots
    from_time = datetime.datetime.strptime('08:00', '%H:%M').time()
    to_time = datetime.datetime.strptime('16:00', '%H:%M').time()
    df_sheet_workday = df_sheet_data.loc[(df_sheet_data['start_time'] >= from_time) & (df_sheet_data['end_time'] <= to_time), df_sheet_data.columns]
    
    if len(df_sheet_workday) > 0:
        start_time = pd.to_datetime(df_sheet_workday['start_time'].astype(str))
        end_time = pd.to_datetime(df_sheet_workday['end_time'].astype(str))
        downtime = end_time - start_time
        
        #Adding a downtime column in dataframe
        df_sheet_workday['Downtime'] = downtime / datetime.timedelta(minutes=1)
    else:
        df_sheet_workday['Downtime'] = 0
    
    #Calculate data for 24 hour
    start_time = pd.to_datetime(df_sheet_data['start_time'].astype(str))
    end_time = pd.to_datetime(df_sheet_data['end_time'].astype(str))
    downtime = end_time - start_time
    
    #Adding a downtime column in dataframe
    df_sheet_data['Downtime'] = downtime / datetime.timedelta(minutes=1)     
    
    return df_sheet_data, df_sheet_workday

#calculation of uptime and dowtime statisitics
def calculate_statistics(df_sheet_data, df_sheet_workday, environments, no_of_days):
    #dictionary to collect statistics
    stats_dict = defaultdict(list)
    stats_dict_workday = defaultdict(list)
       
    
    #Calculate the statistics for each env
    for env in environments:
        if env not in stats_dict:
            pct_planned_downtime = (df_sheet_data[df_sheet_data['Environment'] == env][df_sheet_data['Planned'] == 'Yes']['Downtime'].sum()/4320) *100 
            pct_unplanned_downtime = (df_sheet_data[df_sheet_data['Environment'] == env][df_sheet_data['Planned'] == 'No']['Downtime'].sum()/4320)* 100 
            pct_uptime = 100 - pct_planned_downtime - pct_unplanned_downtime
            summary = "\n".join(list(df_sheet_data[df_sheet_data['Environment'] == env]['summary']))
            stats_dict[env] = [pct_uptime, pct_planned_downtime, pct_unplanned_downtime, summary]
        else:
            print("{0} Environment already exists".format(env))
        
        if env not in stats_dict_workday:
            if len(df_sheet_workday) == 0:
                stats_dict_workday[env] = [100, 0, 0,""]
            else:
                pct_planned_downtime = (df_sheet_workday[df_sheet_workday['Environment'] == env][df_sheet_workday['Planned'] == 'Yes']['Downtime'].sum()/4320) *100 
                pct_unplanned_downtime = (df_sheet_workday[df_sheet_workday['Environment'] == env][df_sheet_workday['Planned'] == 'No']['Downtime'].sum()/4320)* 100 
                pct_uptime = 100 - pct_planned_downtime - pct_unplanned_downtime
                summary = "\n".join(list(df_sheet_workday[df_sheet_workday['Environment'] == env]['summary']))
                stats_dict_workday[env] = [pct_uptime, pct_planned_downtime, pct_unplanned_downtime, summary]              

    return stats_dict, stats_dict_workday

#Function to generate plots and dump it to a PDF 
def generate_plots(labels, colors, stats_dict,stats_dict_workday, environments, min_date, max_date):

    with PdfPages(r'Charts.pdf') as export_pdf:
        rows = 2
        cols = len(environments)
        fig, ax = plt.subplots(rows, cols)
        title = "Workday"
        chart_data = stats_dict_workday
        print(chart_data)
        for row in range(rows):
            for col, env in enumerate(environments):
                ax[row, col].pie(chart_data[env][:3], colors=colors, shadow = False, startangle=90)
                ax[row, col].set_title('{0} {1}'.format(env, title), fontsize=14, fontweight="bold")
                ax[row, col].axis('equal')
                ax[row, col].legend(labels=[f'{x} {np.round(y, 2)}%' for x, y in zip(labels, chart_data[env][:3])], 
                        bbox_to_anchor=(1, 0))
                fig = plt.gcf()
                fig.set_size_inches(15, 10)
                circle = plt.Circle(xy=(0,0), radius=0.70, facecolor='white')
                ax[row, col].add_patch(circle)
            title = "24 Hour"
            chart_data = stats_dict
        
        fig.suptitle("Dev and Clone Environment Availability (Date Range: {0} to {1})".format(min_date, max_date),fontsize=18, fontweight="bold")
        fig.tight_layout() 
        fig.subplots_adjust(top=0.88)
        export_pdf.savefig()
        plt.show()
        plt.close()

 
if __name__ == "__main__":
    print("Start of program\n")
    report_dates, input_file = fetch_cmdline_options()
    df_sheet_data, min_date, max_date, no_of_days = fetch_input_data(report_dates, input_file)
    df_sheet_data, df_sheet_workday = calculate_downtime(df_sheet_data)

    #Intializing the environments 
    environments = ['Development', 'Clone pre_prod','Clone UAT']
    stats_dict, stats_dict_workday = calculate_statistics(df_sheet_data, df_sheet_workday, environments, no_of_days)
    labels= ['Uptime', 'Planned Downtime', 'Unplanned Downtime']
    colors=['green', 'orange', 'red']
    generate_plots(labels, colors, stats_dict, stats_dict_workday, environments, min_date, max_date)
    print("\nEnd of Program")