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
    parser.add_argument('--historical_start_date', dest='historical_start_date', required= True,
                        help='Enter the historical start date in format mm/dd/yyyy')    
    parser.add_argument('--holiday_file', dest='holiday_file', default='holiday.txt',
                        help='provide file containing holidays in format mm/dd/yyyy')
    parser.add_argument('--input_file', dest='input_file', default= 'task_env_report.xlsx',
                        help='provide input excel file containing env availability data')
    #parser.add_argument('--historical_input_file', dest='historical_input_file', default= 'historical_dataset.xlsx',
    #                   help='provide historical excel file containing env availability data')                         
                      
    args = parser.parse_args()

    # TODO: make below multi line
    start_date, end_date, historical_start_date, input_file, holiday_file =  \
                    args.start_date, args.end_date, args.historical_start_date, args.input_file,    \
                    args.holiday_file
    
    #Converting start and end date to datetime objects
    start_date = datetime.datetime.strptime(start_date, '%m/%d/%Y')
    end_date = datetime.datetime.strptime(end_date, '%m/%d/%Y')
    historical_start_date = datetime.datetime.strptime(historical_start_date, '%m/%d/%Y')
    
    #Initial input checks
    if start_date > end_date:
        print("Start Date is greater than end date, therefore, exiting the program")
        # TODO: failures should exit with -1
        exit(-1)
    
    if historical_start_date > end_date:
        print("historical start date is greater than end date, therefore, exiting the program")
        exit(-1)

    # TODO: check if you have read access to the files
    if not os.path.isfile(input_file):
        print("Input File :{0} doesn't exist".format(input_file))
        exit(-1)
    if not os.path.isfile(holiday_file):
        print("Holiday File :{0} doesn't exist".format(holiday_file))
        exit(-1)
    
    return start_date, end_date, historical_start_date, input_file, holiday_file
 
#Fetching data from the excel
def fetch_input_data(input_file):
    df_input_data = pd.read_excel(input_file, sheet_name='env_data')  
    df_input_data['start_time'] = [d.time() for d in df_input_data['start_date']]
    df_input_data['end_time'] = [d.time() for d in df_input_data['end_date']]
   
    return df_input_data

#Split the downtimes for different dates in case the time exceeds 24 hours
# TODO: use same variable format - snake cases preferred (underscore separated)
def split_date(environment,category, planned, start_date, end_date, summary, start_time, end_time):
    #Same day case                                                                                 
    if start_date == end_date:                                                                 
        return [(environment, category, planned, start_date, end_date,  summary, start_time, end_time)]                                                                     
    # TODO: change variable name - and simplify the below code - if its hardcoding, just hard code and add a comment why
    #stop_split = (datetime.datetime.combine(datetime.date(1,1,1),start_time.replace(hour=0, minute=0, second=0)) + datetime.timedelta(days=1)).time()

    midnight_datetime = datetime.datetime.combine(start_date, datetime.datetime.strptime('23:59:00', '%H:%M:%S').time())
    

    #print(midnight_time, type(start_date), start_date, type(start_date + datetime.timedelta(days=1)), start_date + datetime.timedelta(days=1))
    # TODO: the below code wont work if end date is 2 days or more ahead of start date
    return [(environment, category, planned, start_date,start_date, summary, start_time, midnight_datetime.time())] \
              + split_date(environment,category, planned, start_date + datetime.timedelta(days=1), end_date,summary,\
                (midnight_datetime + datetime.timedelta(minutes=1)).time(), end_time)

#Calculating the dates for which report needs to be generated 
def calculate_reporting_date_range(start_date, end_date, holiday_file):
    report_dates_24hr = pd.date_range(start_date, end_date)
    report_dates_workday = pd.bdate_range(start_date, end_date)
    
    #Excluding the holidays for workday plot
    # use context manager to read the files (with syntax) - same for input file
    with open(holiday_file) as file:
      dates_to_be_excluded = file.read().strip().split("\n")

    for dt in dates_to_be_excluded:
        dt = datetime.datetime.strptime(dt, '%m/%d/%Y')
        if dt >= start_date and dt <= end_date:
          report_dates_workday = report_dates_workday.drop(dt)
    #holiday_file.close()
   
    return report_dates_24hr, report_dates_workday
    
#Filtering input data based on dates and timings
def filter_input_data(df_weekly_data, start_date, end_date, holiday_file):
    splitted_dates = [
        elt for _, row in df_weekly_data.iterrows() for elt in split_date(row["Environment"],row["Category"], row["Planned"], row["start_date"].date(), row["end_date"].date(), row["summary"], row["start_time"], row["end_time"])
    ]
    #Here, we have data for each individual date now
    filtered_dataframe = pd.DataFrame(splitted_dates, columns=list(df_weekly_data.columns))
    #print(filtered_dataframe)
    #Filter out the rows for which we need the report for
    report_dates_24hr, report_dates_workday = calculate_reporting_date_range(start_date, end_date, holiday_file)
   
    filtered_dataframe_24hr = filtered_dataframe.loc[(filtered_dataframe['start_date'].isin(report_dates_24hr.date))]
    filtered_dataframe_workday = filtered_dataframe.loc[(filtered_dataframe['start_date'].isin(report_dates_workday.date))]   

    return filtered_dataframe_24hr, filtered_dataframe_workday

def calculate_no_of_days(filtered_dataframe_24hr, filtered_dataframe_workday):
      # TODO: If this is number of dates, change variable to no_of_xyz
    no_of_days_24hr =  7 #len(filtered_dataframe_24hr['start_date'].unique())
    no_of_days_workday =  5 #len(filtered_dataframe_workday['start_date'].unique())
 
    # TODO: make the function name in sync with what is hapenning inside the function
    return no_of_days_24hr, no_of_days_workday

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

    return filtered_dataframe

def calculate_24hr_downtime(filtered_dataframe_24hr):
    start_time = pd.to_datetime(filtered_dataframe_24hr['start_time'].astype(str))
    end_time = pd.to_datetime(filtered_dataframe_24hr['end_time'].astype(str))
    downtime = (end_time - start_time)
    filtered_dataframe_24hr['downtime'] = downtime / datetime.timedelta(minutes=1)

    #required if the entries for end_date are 00:00:00 next day in excel (then downtime will be zero for them)
    filtered_dataframe_24hr = filtered_dataframe_24hr[filtered_dataframe_24hr['downtime'] != 0]

    return filtered_dataframe_24hr
    

#Calculation of downtime-24 hr and workday
def calculate_downtime(filtered_dataframe_24hr, filtered_dataframe_workday):
   
    #Calculate data for 24 hour
    filtered_dataframe_24hr = calculate_24hr_downtime(filtered_dataframe_24hr)   
    
    #Calculate downtime data for workday plots
    filtered_dataframe_workday = calculate_workday_downtime(filtered_dataframe_workday)
    
    return filtered_dataframe_24hr, filtered_dataframe_workday

#calculation of uptime and dowtime statisitics for plotting
def calculate_statistics(filtered_dataframe_24hr, filtered_dataframe_workday, environments, no_of_days_24hr, no_of_days_workday):
    #dictionary to collect statistics
    stats_dict =defaultdict(list)
    stats_dict_workday = defaultdict(list)
    planned_summary_stats =defaultdict(list)
    unplanned_summary_stats =defaultdict(list)
    #print(filtered_dataframe_24hr, filtered_dataframe_workday)
    #Calculate the statistics for each env
    for env in environments:
        if env not in stats_dict: 
            pct_planned_downtime = (filtered_dataframe_24hr[(filtered_dataframe_24hr['Environment'] == env) 
                                    & (filtered_dataframe_24hr['Planned'] == 'Yes')]['downtime'].sum()/(1440*no_of_days_24hr)) *100 
            pct_unplanned_downtime = (filtered_dataframe_24hr[(filtered_dataframe_24hr['Environment'] == env)
                                    & (filtered_dataframe_24hr['Planned'] == 'No')]['downtime'].sum()/(1440*no_of_days_24hr))* 100 
            pct_uptime = 100 - pct_planned_downtime - pct_unplanned_downtime
            planned_summary = set(filtered_dataframe_24hr[(filtered_dataframe_24hr['Environment'] == env) 
                                    & (filtered_dataframe_24hr['Planned'] == 'Yes')]['summary'])
            unplanned_summary = set(filtered_dataframe_24hr[(filtered_dataframe_24hr['Environment'] == env) 
                                    & (filtered_dataframe_24hr['Planned'] == 'No')]['summary'])
            planned_summary_stats[env].append(planned_summary)
            unplanned_summary_stats[env].append(unplanned_summary)
            stats_dict[env] = [pct_uptime, pct_planned_downtime, pct_unplanned_downtime]
        else:
            print("{0} Environment already exists".format(env))
        
        if env not in stats_dict_workday:
            pct_planned_downtime = (filtered_dataframe_workday[(filtered_dataframe_workday['Environment'] == env) 
                                    & (filtered_dataframe_workday['Planned'] == 'Yes')]['downtime_workday'].sum()/(540*no_of_days_workday)) *100 
            pct_unplanned_downtime = (filtered_dataframe_workday[(filtered_dataframe_workday['Environment'] == env) 
                                    & (filtered_dataframe_workday['Planned'] == 'No')]['downtime_workday'].sum()/(540*no_of_days_workday))* 100 
            pct_uptime = 100 - pct_planned_downtime - pct_unplanned_downtime
            planned_summary = set(filtered_dataframe_workday[(filtered_dataframe_workday['Environment'] == env) 
                                    & (filtered_dataframe_workday['Planned'] == 'Yes')]['summary'])
            unplanned_summary = set(filtered_dataframe_workday[(filtered_dataframe_workday['Environment'] == env) 
                                    & (filtered_dataframe_workday['Planned'] == 'No')]['summary'])
            
            stats_dict_workday[env] = [pct_uptime, pct_planned_downtime, pct_unplanned_downtime]
            
            planned_summary_stats[env].append(planned_summary)
            unplanned_summary_stats[env].append(unplanned_summary)

    return stats_dict, stats_dict_workday, planned_summary_stats, unplanned_summary_stats
    
#Function to calculate stats for historical days
def calculate_statistics_historical(hist_df_24hr, hist_df_workday, environments):
    columns = ['Environment','Date', 'Uptime', 'Planned', 'Unplanned'] 
    
    #dictionary to collect statistics
    hist_dict = defaultdict(list)
    hist_dict_workday = defaultdict(list)
    unplanned_summary_stats = {}
    
    #Calculate the statistics for 24 hr
    hist_df_24hr['start_date']  = pd.to_datetime(hist_df_24hr['start_date'])
    hist_df_24hr = hist_df_24hr.groupby(['Environment','Planned','summary',
                                         pd.Grouper(key='start_date', freq='W-SUN')])['downtime'].sum().reset_index().sort_values('start_date')
    hist_df_24hr['start_date']  = hist_df_24hr['start_date'] - datetime.timedelta(days=6)
    hist_df_24hr['downtime'] = hist_df_24hr['downtime']/(1440*7) * 100
    for d in hist_df_24hr['start_date'].unique():
        for env in environments:
            planned = list(hist_df_24hr[(hist_df_24hr['start_date'] == d) 
                        & (hist_df_24hr['Environment'] == env) 
                        & (hist_df_24hr['Planned'] == 'Yes')]['downtime'])
            unplanned = list(hist_df_24hr[(hist_df_24hr['start_date'] == d) 
                        & (hist_df_24hr['Environment'] == env)
                        & (hist_df_24hr['Planned'] == 'No')]['downtime'])
            planned = round(planned[0], 2) if planned else 0
            unplanned = round(unplanned[0], 2) if unplanned else 0
            uptime = 100 - planned - unplanned

            unplanned_summary = (hist_df_24hr[(hist_df_24hr['Environment'] == env )
                                & (hist_df_24hr['Planned'] == 'No')][['start_date','summary']]).groupby(['start_date'], as_index=False).agg(lambda x: set(x))
            #print(unplanned_summary)
            hist_dict[(d,env)]= [env, pd.to_datetime(d).date(), uptime, planned, unplanned]
            unplanned_summary_stats[env] = unplanned_summary

    hist_plot_data_24hr = pd.DataFrame(hist_dict.values(), columns=columns)
    
    #Calculate the stats for workday 
    hist_df_workday['start_date']  = pd.to_datetime(hist_df_workday['start_date'])
    hist_df_workday = hist_df_workday.groupby(['Environment','Planned',
                                               pd.Grouper(key='start_date', freq='W-SUN')])['downtime_workday'].sum().reset_index().sort_values('start_date')
    hist_df_workday['start_date']  = hist_df_workday['start_date'] - datetime.timedelta(days=6)
    hist_df_workday['downtime_workday'] = hist_df_workday['downtime_workday']/(540*5) * 100
        
    for d in hist_df_workday['start_date'].unique():
        for env in environments:
            planned = list(hist_df_workday[(hist_df_workday['start_date'] == d) 
                        & (hist_df_workday['Environment'] == env) 
                        & (hist_df_workday['Planned'] == 'Yes')]['downtime_workday'])
            unplanned = list(hist_df_workday[(hist_df_workday['start_date'] == d) 
                        & (hist_df_workday['Environment'] == env) 
                        & (hist_df_workday['Planned'] == 'No')]['downtime_workday'])
            planned = round(planned[0], 2) if planned else 0
            unplanned = round(unplanned[0], 2) if unplanned else 0
            uptime = 100 - planned - unplanned
            hist_dict_workday[(d,env)]= [env, pd.to_datetime(d).date(), uptime, planned, unplanned]
            
    hist_plot_data_workday = pd.DataFrame(hist_dict_workday.values(), columns=columns)

    return hist_plot_data_24hr, hist_plot_data_workday, unplanned_summary_stats

#Generate historical plots
def generate_historical_plots(labels, colors, hist_plot_data_24hr, hist_plot_data_workday, unplanned_hist_summary_stats, env):
    rows = 2
    cols = 2
    fig1, ax1 = plt.subplots(rows, 2)
    text_content = []
    chart_data = hist_plot_data_24hr
    summary_title = 'Unplanned Downtime'
    chart_title = "Historical {0} 24hr".format(env)
    for row in range(rows):
        for col in range(cols):
            if row >=1:
                ax1[row, col].axis('equal')
                ax1[row, col].axis('off')
                if col >=1:
                    break
                text_content = []
                ax1[row, col].set_title('{0} {1}'.format(env, summary_title),x= 0.3 ,\
                                        y = 0.7,  fontsize=14, fontweight="bold")
                temp_df = unplanned_hist_summary_stats[env][['start_date','summary']]

                #Printing the summary contents of unplanned downtime
                for i in range(len(temp_df)):
                    l = r"$\bf{" + "{0}:".format(temp_df['start_date'][i].date()) \
                        + "}$" +"\n* {0}\n".format("\n* ".join(temp_df['summary'][i]))
                    text_content.append(l)
               
                txt1 = ax1[row, col].text(0.05, 0.70, "".join(text_content), \
                                          ha='left', va = 'top', fontsize=10,\
                                          rotation=0, wrap=True,             \
                                          transform= ax1[row, col].transAxes)
                txt1._get_wrap_line_width = lambda : 800                
            else:
                ax1[row,col].set_title('{0}'.format(chart_title), fontsize=10, \
                                       fontweight="bold")
                chart_data[chart_data['Environment'] == env].plot.bar(ax=ax1[row,col],\
                                                                      x= 'Date',\
                                                                      color = colors,\
                                                                      stacked= True,\
                                                                      legend = False,\
                                                                      width = 0.3,   \
                                                                      rot = 40)
                ax1[row,col].legend(labels, bbox_to_anchor=(1,-0.25),ncol =3)
                ax1[row,col].grid(axis='y')
                ax1[row,col].set_axisbelow(True)
                ax1[row,col].set_ylabel("Percentage (%)", fontsize=10, fontweight="bold")
                ax1[row,col].set_xlabel("WeekOf", fontsize=10, fontweight="bold")            
                fig1 = plt.gcf()
                fig1.set_size_inches(15,12)
                rec = plt.Rectangle((-0.5, 0),  15, 112, fill=False, lw=3, zorder=20)
                ax1[row,col].add_patch(rec)
        
            chart_data = hist_plot_data_workday
            chart_title = "Historical {0} Workday".format(env)
    fig1.suptitle("Historical Environment Availability : {0}".format(env),fontsize=14, fontweight="bold")
    fig1.subplots_adjust(top=0.93)
    
    return fig1


#Function to generate plots and dump it to a PDF 
def generate_plots(labels, colors, stats_dict,stats_dict_workday, planned_summary_stats, unplanned_summary_stats, environments, start_date, end_date):
    rows = 4
    cols = len(environments)
    fig, ax = plt.subplots(rows, cols)
    title = "Workday"
    summary_title = 'Unplanned Downtime'
    summary_data = unplanned_summary_stats
    chart_data = stats_dict_workday
    
    for row in range(rows - 2):
        for col, env in enumerate(environments):
            #chart_data[env][:3] = [round(elem) for elem in chart_data[env][:3] ]
            #print(chart_data[env][:3] )
            ax[row, col].pie(chart_data[env][:3], colors=colors, shadow = False,  \
                             startangle=90, autopct=lambda p: '{:.2f}%'.          \
                             format(round(p,2)) if p > 0 else '',
                             labeldistance =1.5)
            ax[row, col].axis('equal')
            ax[row, col].text(0, 0, '{0} {1}'.format(env, title), va = 'center',  \
                              ha = 'center', fontsize=10, fontweight="bold")
            fig = plt.gcf()
            fig.set_size_inches(15, 20)
            circle = plt.Circle(xy=(0,0), radius=0.80, facecolor='white')
            rec = plt.Rectangle((-1.1, -1.1),  2.2, 2.19, fill=False, lw=3, zorder=100)
            ax[row, col].add_patch(rec)
            ax[row, col].add_patch(circle)              
        title = "24 Hr"
        chart_data = stats_dict
    
    for row in range(2, rows):
        for col, env in enumerate(environments):                  
                ax[row, col].set_title('{0} {1}'.format(env, summary_title), fontsize=12, fontweight="bold")
                ax[row, col].axis('equal')
                ax[row, col].axis('off')
                txt = ax[row, col].text(0.05, 1, "{1} {0}".format("\n* ".join(summary_data[env][0]), \
                                                                  "*" if summary_data[env][0] else ""),\
                                        ha='left', 
                                        va = 'top', 
                                        fontsize=12, 
                                        rotation=0, 
                                        wrap=True,
                                        transform=ax[row, col].transAxes)
                txt._get_wrap_line_width = lambda : 250
                fig = plt.gcf()
                fig.set_size_inches(15, 20)
        summary_title = "Planned Downtime"
        summary_data = planned_summary_stats

    fig.suptitle("Dev and Clone Environment Availability (Date Range: {0} to {1})". \
                 format(start_date.date(),
                        end_date.date()),
                        fontsize=14, 
                        fontweight="bold")
    fig.legend(labels=[f'{x}' for x in labels],bbox_to_anchor=(0.81, 0.96), \
               columnspacing = 10,prop={'size': 12}, ncol = 3)
    fig.tight_layout() 
    fig.subplots_adjust(top=0.93)
    return fig
        
#Add pages to PDF file 
def generate_pdf(pages):
    with PdfPages(r'Env_Availability_report.pdf') as export_pdf:
        for page in pages:
            export_pdf.savefig(page)
            #plt.show()
            plt.close()

#Main Program            
if __name__ == "__main__":
    print("Start of program\n")
    environments = ['Development', 'Clone Pre Prod','Clone UAT']
    labels= ['Uptime', 'Planned Downtime', 'Unplanned Downtime']
    colors=['green', 'orange', 'red']
    #input_file = '/content/drive/MyDrive/task_env_report_v1.0.xlsx'

    #start_date = '07/30/2022' 
    #end_date = '08/05/2022'
    #start_date = datetime.datetime.strptime(start_date, '%m/%d/%Y')
    #end_date = datetime.datetime.strptime(end_date, '%m/%d/%Y')
    
    #input_file = '/content/drive/MyDrive/task_env_report_v1.0.xlsx'
    #holiday_file = '/content/drive/MyDrive/holiday.txt'
    #historical_start_date = '06/20/2022'
    #historical_start_date = datetime.datetime.strptime(historical_start_date, '%m/%d/%Y')
    start_date, end_date, historical_start_date, input_file, holiday_file = fetch_cmdline_options()
    
    #start_date, end_date, historical_start_date, input_file, historical_input_file, holiday_file = fetch_cmdline_options()
    
    #Weekly Plotss
    input_data = fetch_input_data(input_file)
    filtered_dataframe_24hr, filtered_dataframe_workday = filter_input_data(input_data, start_date, end_date, holiday_file)
    no_of_days_24hr, no_of_days_workday = calculate_no_of_days(filtered_dataframe_24hr, filtered_dataframe_workday)
    
    filtered_dataframe_24hr, filtered_dataframe_workday = calculate_downtime(filtered_dataframe_24hr, filtered_dataframe_workday)
    stats_dict, stats_dict_workday, planned_summary_stats, unplanned_summary_stats = calculate_statistics(filtered_dataframe_24hr, filtered_dataframe_workday, environments, no_of_days_24hr, no_of_days_workday)
    pages= [generate_plots(labels, colors, stats_dict,stats_dict_workday, planned_summary_stats, unplanned_summary_stats, environments, start_date, end_date)]
    
    
    #Historical Plots
    hist_df_24hr, hist_df_workday = filter_input_data(input_data, historical_start_date, end_date, holiday_file)
    hist_df_24hr, hist_df_workday = calculate_downtime(hist_df_24hr, hist_df_workday)
    hist_plot_data_24hr, hist_plot_data_workday, unplanned_hist_summary_stats = calculate_statistics_historical(hist_df_24hr, hist_df_workday, environments)
    #Add pages to final PDF file.
    for env in environments:
        page = generate_historical_plots(labels, colors, hist_plot_data_24hr, hist_plot_data_workday, unplanned_hist_summary_stats,env)
        pages.append(page)
        
    generate_pdf(pages)
    
    print("\nEnd of Program")