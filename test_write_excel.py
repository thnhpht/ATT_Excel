import openpyxl
import os
import sys
from zk import ZK, const
import datetime
import calendar

CWD = os.path.dirname(os.path.realpath(__file__))
ROOT_DIR = os.path.dirname(CWD)
sys.path.append(ROOT_DIR)

conn = None
zk = ZK('192.168.1.201', port=4370)
conn = zk.connect()

def get_users():
    arr_users = []
    users = conn.get_users()
    for user in users:
        arr_user = []
        if user.privilege != const.USER_ADMIN:
            arr_user.append(user.user_id)
            arr_user.append(user.name)
            arr_users.append(arr_user)
    return arr_users

def get_date_time(month, year):
    arr_datetime = []
    num_days = calendar.monthrange(year, month)[1]
    if int(month) < 10:
            month = "0" + str(month)
    for day in range(1, num_days + 1):
        if int(day) < 10:
            day = "0" + str(day)

        datetime = str(day) + "/" + str(month) + "/" + str(year)
        arr_datetime.append(datetime)
    return arr_datetime

def get_day_in_week(s):
    day, month, year = (int(x) for x in s.split("/"))
    day_in_week = datetime.date(year, month, day)
    if day_in_week.strftime("%A") == "Sunday":
        return "Cn"
    return "HC"
     
def get_clock_in(user_id, date):
    attendances = conn.get_attendance()
    day, month, year = (str(x) for x in date.split("/"))
    month_day = month + "-" + day
    for x in attendances:
        x = str(x).replace("<Attendance>: ", "")
        x2 = x.split(":")

        if x2[0].strip() != user_id or month_day not in x2[1]: 
            continue

        x3 = x2[1].strip().split()
        return str(x3[1]) + ":" + str(x2[2])      

def get_clock_out(user_id, date):
    attendances = conn.get_attendance()
    day, month, year = (str(x) for x in date.split("/"))
    month_day = month + "-" + day
    count = 0
    for x in attendances:
        x = str(x).replace("<Attendance>: ", "")
        x2 = x.split(":")

        if x2[0].strip() != user_id or month_day not in x2[1]: 
            continue
        
        count += 1

        if count == 2:
            x3 = x2[1].strip().split()
            return str(x3[1]) + ":" + str(x2[2])     

def output_Excel(input_detail, output_excel_path):
  #Xác định số hàng và cột lớn nhất trong file excel cần tạo
  row = len(input_detail)
  column = len(input_detail[0])

  #Tạo một workbook mới và active nó
  wb = openpyxl.Workbook()
  ws = wb.active
  
  #Dùng vòng lặp for để ghi nội dung từ input_detail vào file Excel
  for i in range(0,row):
    for j in range(0,column):
        try:
            v=input_detail[i][j]
            ws.cell(column=j+1, row=i+1, value=v)
        except:
            break

  #Lưu lại file Excel
  wb.save(output_excel_path)


input_detail = [["AC-No.", "No.", "Name", "Date", "Timetable", "On duty", "Off duty", "Clock in", "Clock out", "Normal", "Real time", "Late", "Early", "Absent", "OT Time", "Work Time", "Exception", "Must C/In", "Must C/Out", "Department", "NDays", "WeekEnd", "Holiday", "ATT_Time", "NDays_OT", "WeekEnd_OT", "Holiday_OT"]]
month = 12
year = 2022
for i in range(len(get_users())):
    for j in range(len(get_date_time(month, year))):
        input_arr = []
        ac_no = get_users()[i][0]
        no = "AH-" + str(get_users()[i][0])
        name = get_users()[i][1]
        date = get_date_time(month, year)[j]
        time_table = get_day_in_week(get_date_time(month, year)[j])
        on_duty = "08:00"
        off_duty = "18:00"
        clock_in = get_clock_in(ac_no, date)
        clock_out = get_clock_out(ac_no, date)


        input_arr.append(ac_no)
        input_arr.append(no)
        input_arr.append(name)
        input_arr.append(date)
        input_arr.append(time_table)
        input_arr.append(on_duty)
        input_arr.append(off_duty)
        input_arr.append(clock_in)
        input_arr.append(clock_out)

        input_detail.append(input_arr)

output_excel_path= 'D:/Test/test.xlsx'
output_Excel(input_detail,output_excel_path)
