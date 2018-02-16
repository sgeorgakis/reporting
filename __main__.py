from service_manager import *
from workbook import *
from datetime import date

date_format = '%d.%m.%Y'

#first_day_of_month = time.strptime(date(d.year, d.month, 1), date_format)
#current_day_of_month = time.strptime(date.today(), date_format)
first_day_of_month = "01/07/2016 00:00:00"
current_day_of_month = "26/07/2016 23:59:59"

manager = ServiceManager(first_day_of_month, current_day_of_month, 'local')

accumulated_data = manager.get_reporting_data()

reporting_data = ["", "", "", "", "", "", "", ""]

for data_list in accumulated_data:
    if "ALL_CLEAN_WITHOUT_ALLOWED" in data_list:
        reporting_data[1] = data_list
    elif "Unknown" in data_list:
        reporting_data[2] = data_list
    elif "Tarif_1" in data_list:
        reporting_data[3] = data_list
    elif "Tarif_2" in data_list:
        reporting_data[4] = data_list
    elif "Tarif_3" in data_list:
        reporting_data[5] = data_list
    elif "Tarif_4" in data_list:
        reporting_data[6] = data_list
    elif "Tarif_5" in data_list:
        reporting_data[7] = data_list

wb = Workbook('2017_03_13_DailyReport.XLSX', 8, 10)
wb.open_sheet(raw_data_sheet)
wb.add_data(reporting_data)
wb.update_raw_data_chart()
wb.create_line_chart()

wb = Workbook('2017_03_13_DailyReportHybrid.xlsx', 1, 4)
wb.open_sheet(raw_data_sheet)
wb.add_data(reporting_data)
wb.update_raw_data_chart()
wb.create_line_chart()