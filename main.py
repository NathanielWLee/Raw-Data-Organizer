import calendar
import plotly
import plotly.express as px
import xlsxwriter
import pandas as pd
import numpy as np
import csv
import tkinter as tk
from tkinter import filedialog
import csv


def main():
    # prompt user to select file
    # open and read file and store inside df
    # df = pd.read_csv('/content/drive/MyDrive/proj/test.csv') # (only useful in colab)

    root = tk.Tk()
    root.withdraw()

    print("Please select a file")
    file_path = filedialog.askopenfilename()
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path, encoding='cp1252')
    else:
        print('Please use Excel or CSV files only!!!')

    with open('excelTest.csv', 'w') as file:
        writer = csv.writer(file)

        df2 = df[['Service', 'Request-Type', 'Fulfillment End Date']]  # filter out only the columns that we need
        df3 = df2.sort_values(['Service', 'Request-Type'], ascending=(True,
                                                                      True))
        # sort based on service then request type for visual purposed

        df3[['Fulfillment End Date']] = df3[['Fulfillment End Date']].apply(
            pd.to_datetime)  # converts to datetime format

        df3['month'] = (
            df3['Fulfillment End Date'].dt.month)  # adds new column called "month" with its corresponding month
        df3 = df3.sort_values(['Service', 'Request-Type', 'month'], ascending=(True, True, True))
        df3['month'] = df3['month'].apply(lambda x: calendar.month_name[x])

        with pd.option_context('display.max_rows', None,
                               'display.max_columns', None,
                               'display.precision', 3,
                               ):
            # total for each service
            print(df3.groupby(['Service'])['Service'].count())
            (df3.groupby(['Service'])['Service'].count()).to_csv("excelTest.csv")

            print("\n")

            # total for each request type in each service
            print(df3.groupby(['Service', 'Request-Type'])['Request-Type'].count())
            (df3.groupby(['Service', 'Request-Type'])['Request-Type'].count()).to_csv("excelTest.csv",
                                                                                      mode='a')
            print("\n")
            # writer.writerow("")

            # total for each request type in each service per month
            print(df3.groupby(['Service', 'Request-Type', 'month'], sort=False)['month'].count())
            (df3.groupby(['Service', 'Request-Type', 'month'], sort=False)['month'].count()).to_csv(
                "excelTest.csv",
                mode='a')
            print("\n")
            # writer.writerow("")

            # total number of services completed within each month (Graph this one)
            print(df3.groupby(['Service', 'month'], sort=False)['month'].count())
            print(df3.groupby(['month'], sort=False)['month'].count())
            df4 = (df3.groupby(['month'], sort=False)['month'].count())
            print(df4.sum())
            (df3.groupby(['Service', 'month'], sort=False)['month'].count()).to_csv("excelTest.csv", mode='a')
            (df3.groupby(['month'], sort=False)['month'].count()).to_csv("excelTest.csv", mode='a')
            pd.DataFrame([[df4.sum()]], columns=list('a')).to_csv("excelTest.csv", mode='a')

            # workbook = xlsxwriter.workbook('excelTest.xlsx')
            # worksheet = workbook.add_worksheet()
            # bold = workbook.add_format({'bold': 1})
            # chart1 = workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})
            # chart1.add_series({
            #
            # })

            # df4['gtPerMonth'] = (df3.groupby(['month'], sort=False)['month'].count()).to_csv("excelTest.xlsx",
            # mode='a')

            # combined = pd.DataFrame()
            # combined = combined.append(df4, ignore_index=True)
            #
            # #Create bar chart
            # chart = px.bar(combined, x='month', y='month')
            # #Save Bar Chart and export to Excel
            # plotly.offline.plot(chart, filename='TestBarChart.xlsx')


if __name__ == "__main__":
    main()
