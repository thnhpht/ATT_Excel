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
        attendances += conn.get_attendance()
    except:
        print("Kiểm tra lại máy chấm công số seri:", x["id"])
        exit()

# Lấy thông tin tất cả người dùng trong máy
users = conn.get_users()

# Lấy số phút cần trừ ra trước giờ chấm công
minute_must_ci = config["minute"]

# Lấy on duty, off duty
list_on_duty = []
list_off_duty = []

for shift in config["shift"]:
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
        # Đưa dữ liệu không nằm trong khoảng thời gian vào mảng để xóa
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
    res = []
    current = date_from

    while current <= date_to:
        res.append(current)
        current += timedelta(days=1)
    return res

# Hàm kiểm tra trong khoảng thời gian
def datetime_in_range(start, end, current):
    return start <= current <= end

# Hàm lấy timetable
def check_sunday(current):
    if current.strftime("%A") == "Sunday":
        return "CN"
    return "HC"

# Hàm lấy giờ chấm công vào
def get_clock_in(current):
    global list_datetime
    # Lấy thời gian on duty, off duty
    time_on_duty = datetime.time(int(on_duty.split(":")[0]), int(on_duty.split(":")[1]))
    hour_on_duty = int(time_on_duty.hour)
    minute_on_duty = int(time_on_duty.minute)
    time_off_duty = datetime.time(int(off_duty.split(":")[0]), int(off_duty.split(":")[1]))
    hour_off_duty = int(time_off_duty.hour)
    minute_off_duty = int(time_off_duty.minute)
    # Lấy thời gian bắt đầu và kết thúc
    start = current + timedelta(hours=hour_on_duty, minutes=minute_on_duty) - timedelta(minutes=minute_must_ci)
    # Kiểm tra nếu off duty < on duty thì cộng 1 ngày
    if time_off_duty < time_on_duty:
        end = current + timedelta(hours=hour_off_duty, minutes=minute_off_duty) + timedelta(days=1)
    else:
        end = current + timedelta(hours=hour_off_duty, minutes=minute_off_duty)
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
    # Lấy thời gian on duty, off duty
    hour_on_duty = int(on_duty.split(":")[0])
    minute_on_duty = int(on_duty.split(":")[1])
    # Lấy thời gian bắt đầu và kết thúc
    start = current + timedelta(hours=hour_on_duty, minutes=minute_on_duty)
    end = current + timedelta(hours=hour_on_duty, minutes=minute_on_duty) + timedelta(days=1) - timedelta(minutes=minute_must_ci, seconds=1)
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
date_from = datetime.datetime(2023, 2, 27)
date_to = datetime.datetime(2023, 3, 1)
date = get_date()

# Lặp lấy thông tin và đưa vào mảng
for i in range(len(users)):
    ac_no = users[i].user_id
    no = users[i].user_id
    name = users[i].name
    list_datetime_ = push_data_into_array(ac_no)

    for j in range(len(date)):
        current = date[j]

        for k in range(len(list_on_duty)):
            row = []
            list_datetime = list_datetime_
            on_duty = list_on_duty[k]
            off_duty = list_off_duty[k]
            time_table = check_sunday(current)
            clock_in = get_clock_in(current)
            clock_out = get_clock_out(current) if clock_in != None else None
         
            row.append(int(ac_no))
            row.append(no)
            row.append(name)
            row.append(current.strftime("%d/%m/%Y"))
            row.append(time_table)
            row.append(on_duty)
            row.append(off_duty)
            row.append(clock_in)
            row.append(clock_out)

            excel.append(row)

# Đường dẫn xuất file
path = 'D:/DuLieuChamCong.xlsx'
# Xuất file
output_Excel(excel, path)
print("Xuất file Excel thành công")

