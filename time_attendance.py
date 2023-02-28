import openpyxl
import json
from zk import ZK, const
import datetime
from datetime import timedelta

# Mở file config
f = open("config.json", encoding="utf-8")
  
# Trả về dạng dictionary
config = json.load(f)

# Lặp và kết nối các máy chấm công từ file config
attendances = []

for x in config["time_attendance"]:
    conn = None
    zk = ZK(x["ip"], port=x["port"])
    try:
        conn = zk.connect()
        attendance = conn.get_attendance()
        attendances += attendance
    except:
        print("Kiểm tra lại máy chấm công số seri:", x["id"])
        exit()

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
minute_must_ci = config["minute"]

# Lấy on duty, off duty
list_on_duty = []
list_off_duty = []

for shift in config["shift"]:
    if shift["shift_code"] == "CS" or shift["shift_code"] == "CC":
        list_on_duty.append(shift["work_start"])
        list_off_duty.append(shift["work_end"])

# Đóng file config
f.close()

# Hàm đưa dữ liệu chấm công vào mảng kiểu datetime theo userid
def push_data_into_array(user_id):
    res = []
    del_list = []
    # Lặp tất cả dữ liệu chấm công
    for attendance in attendances:
        # Xóa dữ liệu không nằm trong khoảng thời gian
        if not datetime_in_range(date_from, date_to, attendance.timestamp):
            del_list.append(attendance)
            continue
        # Nếu đúng userid lấy datetime và đưa vào mảng
        if attendance.user_id == user_id:
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
def output_Excel(input_detail, path):
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
  wb.save(path)

# Tiêu đề excel
header = [["AC-No.", "No.", "Name", "Date", "Timetable", "On duty", "Off duty", "Clock In", "Clock Out", "Normal", "Real time", "Late", "Early", "Absent", "OT Time", "Work Time", "Exception", "Must C/In", "Must C/Out", "Department", "NDays", "WeekEnd", "Holiday", "ATT_Time", "NDays_OT", "WeekEnd_OT", "Holiday_OT"]]
excel = header
# Đưa vào thông tin muốn lấy ngày chấm công
date_from = datetime.datetime(2023, 2, 1)
date_to = datetime.datetime(2023, 3, 1)
date = get_date()

# Lặp lấy thông tin và đưa vào mảng
for i in range(len(list_users)):
    ac_no = list_users[i][0]
    no = "AH-" + list_users[i][0] 
    name = list_users[i][1]
    list_datetime_ = push_data_into_array(ac_no)

    for j in range(len(date)):
        current = date[j]

        for k in range(len(list_on_duty)):
            row = []
            list_datetime = list_datetime_
            on_duty = list_on_duty[k]
            off_duty = list_off_duty[k]
            time_table = get_day_in_week(current)
            clock_in = get_clock_in(current)
            clock_out = get_clock_out(current) if clock_in != None else None
            # normal = None
            # work_time = None
            # real_time = None
            # late = None
            # early = None
            # absent = None
            # ot_time = None
            # exception = None
            # must_clock_in = None 
            # must_clock_out = None
            # department = None
            # n_days = None
            # weekend = None
            # holiday = None
            # att_time = None
            # n_days_ot = None
            # weekend_ot = None
            # holiday_ot = None

            row.append(int(ac_no))
            row.append(no)
            row.append(name)
            row.append(current.strftime("%d/%m/%Y"))
            row.append(time_table)
            row.append(on_duty)
            row.append(off_duty)
            row.append(clock_in)
            row.append(clock_out)

            for i in range(19):
                row.append(None)
            # row.append(normal)
            # row.append(real_time)
            # row.append(late)
            # row.append(early)
            # row.append(absent)
            # row.append(ot_time)
            # row.append(work_time)
            # row.append(exception)
            # row.append(must_clock_in)
            # row.append(must_clock_out)
            # row.append(department)
            # row.append(n_days)
            # row.append(weekend)
            # row.append(holiday)
            # row.append(att_time)
            # row.append(n_days_ot)
            # row.append(weekend_ot)
            # row.append(holiday_ot)

            excel.append(row)

# Đường dẫn xuất file
path = 'D:/DuLieuChamCong.xlsx'
# Xuất file
output_Excel(excel, path)
print("Xuất file Excel thành công")

