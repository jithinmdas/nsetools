"""
Author: Jithin M
Email: jmd@jithinmdas.com
This program is used for equity analysis. It is mainly for indices.
"""

from os import write
import xlsxwriter
import pandas as pd
from nsepy import get_history
from nsepy import get_index_pe_history
from datetime import date


class NseData:
    def __init__(self, symbol, start_date, end_date) -> None:
        self.symbol = symbol
        self.start_date = start_date
        self.end_date = end_date

    def get_price(self):
        data = get_history(self.symbol, self.start_date,
                           self.end_date, index=True)
        print(data)
        return data

    def get_pe_pb_div(self):
        data = get_index_pe_history(
            self.symbol, self.start_date, self.end_date)
        print(data)
        return data


def store_data_xlsx(filename, data, sheetname):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter',
                            date_format='dd mmm yyyy')
    data.to_excel(writer, sheet_name=sheetname)
    workbook = writer.book
    worksheet = writer.sheets[sheetname]
    format1 = workbook.add_format({'num_format': '#,##0.00'})
    worksheet.set_column('A:A', 18, None)
    worksheet.set_column('B:B', None, format1)
    # chart = workbook.add_chart({'type':'line'})
    # chart.add_series({'values':'=Sheet1!$B$2:$B$8'})
    # worksheet.insert_chart('D2', chart)
    writer.save()


def update_mean_till_date(dataframe, column_to_refer, column_to_add):
    for itr in range(1, len(dataframe)):
        sub_mean = pe_pb_div[[column_to_refer]].head(itr).mean(skipna=True)
        result[column_to_add].iat[itr-1] = sub_mean


nse = NseData("NIFTY 50", date(2015, 1, 1), date(2015, 1, 10))

# Get the tables
price = nse.get_price()
pe_pb_div = nse.get_pe_pb_div()

# Concat two tables
result = pd.concat([price, pe_pb_div], axis=1, join="inner")

# Add new columns
result['Average P/E'] = pe_pb_div['P/E']
result['Average P/B'] = pe_pb_div['P/B'].mean(skipna=True)
result['Average Div'] = pe_pb_div['Div Yield'].mean(skipna=True)

# Find mean till date and update the columns
update_mean_till_date(result, 'P/E', 'Average P/E')
update_mean_till_date(result, 'P/B', 'Average P/B')
update_mean_till_date(result, 'Div Yield', 'Average Div')

print(result)

store_data_xlsx('nsedata.xlsx', result, 'NIFTY 50')
