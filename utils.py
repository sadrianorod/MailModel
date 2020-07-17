from config_maildriver import *
from datetime import datetime,timedelta
from bizdays import Calendar
import io
import pandas as pd
from anbima import holidays

def date_after(date,number_days):
    cal252 = Calendar(holidays(), ['Saturday', 'Sunday'])
    _date = datetime.strptime(date,'%d/%m/%Y')
    
    for i in range(number_days):
        _date = cal252.adjust_next(_date+timedelta(days=1))
    
    return _date.strftime("%d/%m/%Y")

def since_date(days):
    """
        param days: Today we use 1
        ptype days: Int
        return: Example '20-Jun-2020'
        rtype: String
    """
    mydate = datetime.now() - timedelta(days=days)
    return mydate.strftime("%d-%b-%Y")


def export_csv(df):
  with io.StringIO() as buffer:
    df.to_csv(buffer,sep=';', index = False)
    return buffer.getvalue()

def export_excel(df,change_format = False):
  with io.BytesIO() as buffer:
    writer = pd.ExcelWriter(buffer,engine='xlsxwriter',date_format=SHEET_DATE_FORMAT)
    
    df.to_excel(writer,index=False,sheet_name=SHEET_NAME)
    
    if change_format:
        workbook  = writer.book
        worksheet = writer.sheets[SHEET_NAME]
        num_format = workbook.add_format({'num_format': SHEET_NUMBER_FORMAT})
        date_format = workbook.add_format({'num_format': SHEET_DATE_FORMAT})
        worksheet.set_column(CELLS_CONTAINING_NUMBER, 18, num_format)
        worksheet.set_column(CELLS_CONTAINING_DATE,18,date_format)
    
    writer.save()
    return buffer.getvalue()
