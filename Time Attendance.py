import openpyxl
import json
from zk import ZK, const
import datetime
from datetime import timedelta

# Mở file config
f = open("D:\Web API\config.json")
  
# Trả về dạng dictionary
data = json.load(f)

# Lặp và kết nối các máy chấm công từ file config
attendances = []

for x in data["time_attendance"]:
    conn = None
    zk = ZK(x["ip"], port=x["port"])
    try:
        conn = zk.connect()
        attendance = conn.get_attendance()
        attendances += attendance
    except:
        print("Khởi động lại máy chấm công serial number:", x["id"])
        exit()
print(attendances)
# Lấy thông tin tất cả người dùng trong máy
list_users = []
users = conn.get_users()
for user in users:
    list_user = []
    if user.privilege != const.USER_ADMIN:
        list_user.append(user.user_id)
        list_user.append(user.name)
        list_users.append(list_user)

# Lấy số phút cần trừ ra trước giờ chấm công
minute_must_ci = data["minute"]

# Đóng file 
f.close()

# Hàm đưa dữ liệu chấm công vào mảng kiểu datetime theo userid
def push_data_into_array(user_id):
    res = []
    del_list = []
    # Lặp tất cả dữ liệu chấm công
    for attendance in attendances:
        # Nếu không đúng userid lặp tiếp
        if user_id != attendance.user_id:
            continue
        # Nếu đúng userid lấy datetime và đưa vào mảng
        date_time = attendance.timestamp
        res.append(date_time)
        # Đưa vào mảng để xóa 
        del_list.append(attendance)
    for x in del_list:
        attendances.remove(x)
    # Trả về mảng datetime được sắp xếp
    return sorted(res)

# Hàm lấy các ngày từ khoảng from to
def get_date():
    arr_datetime = []
    current = date_from
    arr_datetime.append(current) 

    while current < date_to:
        current += timedelta(days=1)
        arr_datetime.append(current)
    return arr_datetime

# Hàm kiểm tra trong khoảng thời gian
def datetime_in_range(start, end, current):
    return start <= current <= end

# Hàm lấy timetable
def get_day_in_week(current):
    day = current.day
    month = current.month
    year = current.year
    day_in_week = datetime.date(year, month, day)
    if day_in_week.strftime("%A") == "Sunday":
        return "Cn"
    return "HC"

# Hàm lấy giờ chấm công vào
def get_clock_in(current):
    global list_datetime
    # Lấy ngày tháng năm 
    day = current.day
    month = current.month
    year = current.year
    # Lấy thời gian on duty, off duty
    time_on_duty = datetime.time(int(on_duty.split(":")[0]), int(on_duty.split(":")[1]))
    time_off_duty = datetime.time(int(off_duty.split(":")[0]), int(off_duty.split(":")[1]))
    hour_on_duty = int(time_on_duty.hour)
    minute_on_duty = int(time_on_duty.minute)
    hour_off_duty = int(time_off_duty.hour)
    minute_off_duty = int(time_off_duty.minute)
    # Lấy thời gian bắt đầu và kết thúc
    start = datetime.datetime(year, month, day, hour_on_duty, minute_on_duty) - timedelta(minutes=minute_must_ci)
    # Kiểm tra nếu off duty < on duty thì cộng 1 ngày
    if time_off_duty < time_on_duty:
        end = datetime.datetime(year, month, day, hour_off_duty, minute_off_duty) + timedelta(days=1)
    else:
        end = datetime.datetime(year, month, day, hour_off_duty, minute_off_duty)
    # Lặp tất cả datetime trong mảng
    for id, current in enumerate(list_datetime):
        # Nếu current đúng trong khoảng thời gian start end trả về current
        if datetime_in_range(start, end, current):
            list_datetime = list_datetime[id+1:]
            return str(current)

# Hàm lấy giờ chấm công ra
def get_clock_out(current):
    global list_datetime
    res = ""
    flag = False
    # Lấy ngày tháng năm 
    day = current.day
    month = current.month
    year = current.year
    # Lấy thời gian on duty, off duty
    time_on_duty = datetime.time(int(on_duty.split(":")[0]), int(on_duty.split(":")[1]))
    hour_on_duty = int(time_on_duty.hour)
    minute_on_duty = int(time_on_duty.minute)
    # Lấy thời gian bắt đầu và kết thúc
    start = datetime.datetime(year, month, day, hour_on_duty, minute_on_duty)
    end = datetime.datetime(year, month, day, hour_on_duty, minute_on_duty) + timedelta(days=1) - timedelta(minutes=minute_must_ci, seconds=1)
    # Lặp tất cả datetime trong mảng
    for id, current in enumerate(list_datetime):
        # Nếu current đúng trong khoảng thời gian start end trả về current
        if datetime_in_range(start, end, current):
            res = current
            flag = True
            continue
        if flag == True:
            list_datetime = list_datetime[id:]
            break
    return str(res)

# Hàm xuất file excel
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

# Tiêu đề excel
input_detail = [["AC-No.", "No.", "Name", "Date", "Timetable", "On duty", "Off duty", "Clock In", "Clock Out", "Normal", "Real time", "Late", "Early", "Absent", "OT Time", "Work Time", "Exception", "Must C/In", "Must C/Out", "Department", "NDays", "WeekEnd", "Holiday", "ATT_Time", "NDays_OT", "WeekEnd_OT", "Holiday_OT"]]

# Đưa vào thông tin muốn lấy ngày chấm công
date_from = datetime.date(2023, 2, 1)
date_to = datetime.date(2023, 2, 21)
date = get_date()
on_duty = "08:30"
off_duty = "17:30"

# Lặp lấy thông tin và đưa vào mảng
for i in range(len(list_users)):
    ac_no = list_users[i][0]
    no = "AH-" + list_users[i][0] 
    name = list_users[i][1]
    list_datetime = push_data_into_array(ac_no)
    for j in range(len(date)):
        input_arr = []
        current = date[j]
        time_table = get_day_in_week(current)
        clock_in = get_clock_in(current)
        clock_out = get_clock_out(current)
        normal = None
        work_time = None
        real_time = None
        late = None
        early = None
        absent = None
        ot_time = None
        exception = None
        must_clock_in = None 
        must_clock_out = None
        department = None
        n_days = None
        weekend = None
        holiday = None
        att_time = None
        n_days_ot = None
        weekend_ot = None
        holiday_ot = None

        input_arr.append(int(ac_no))
        input_arr.append(no)
        input_arr.append(name)
        input_arr.append(current .strftime("%d/%m/%Y"))
        input_arr.append(time_table)
        input_arr.append(on_duty)
        input_arr.append(off_duty)
        input_arr.append(clock_in)
        input_arr.append(clock_out)
        input_arr.append(normal)
        input_arr.append(real_time)
        input_arr.append(late)
        input_arr.append(early)
        input_arr.append(absent)
        input_arr.append(ot_time)
        input_arr.append(work_time)
        input_arr.append(exception)
        input_arr.append(must_clock_in)
        input_arr.append(must_clock_out)
        input_arr.append(department)
        input_arr.append(n_days)
        input_arr.append(weekend)
        input_arr.append(holiday)
        input_arr.append(att_time)
        input_arr.append(n_days_ot)
        input_arr.append(weekend_ot)
        input_arr.append(holiday_ot)

        input_detail.append(input_arr)

# Đường dẫn xuất file
output_excel_path= 'D:/Test/test.xlsx'
# Xuất file
output_Excel(input_detail,output_excel_path)
print("Xuất file Excel thành công")

