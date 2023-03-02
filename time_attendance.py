import openpyxl
import json
from zk import ZK
import datetime
from datetime import timedelta


# Hàm lấy thời gian chấm công của người dùng trong khoảng thời gian theo userid
def get_attendances_of_user(user_id, attendances, date_from, date_to):
    res = []
    del_list = []
    # Lặp và lấy tất cả thời gian chấm công đúng userid nằm trong khoảng thời gian và xóa đi
    for attendance in attendances:
        if not datetime_in_range(date_from, date_to, attendance.timestamp):
            del_list.append(attendance)
            continue
        if attendance.user_id == user_id:
            date_time = attendance.timestamp
            res.append(date_time)
            # Đưa vào mảng để xóa 
            del_list.append(attendance)
    for del_item in del_list:
        attendances.remove(del_item)
    # Trả về thời gian chấm công đã được sắp xếp
    return sorted(res)


# Hàm lấy tất cả các ngày từ khoảng thời gian bắt đầu - kết thúc
def get_date(date_from, date_to):
    res = []
    current = date_from

    while current <= date_to:
        res.append(current)
        current += timedelta(days=1)
    return res


# Hàm kiểm tra ngày đang xét có nằm trong khoảng thời gian, nếu đúng trả về true ngược lại trả về false
def datetime_in_range(start, end, current):
    return start <= current <= end


# Hàm lấy giờ chấm công vào
def get_clock_in(current, on_duty, off_duty, minute_must_ci):
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
    # Lặp và lấy đúng thời gian chấm công vào, không có trả về None
    for i, current in enumerate(list_datetime):
        if datetime_in_range(start, end, current):
            list_datetime = list_datetime[i + 1:]
            return str(current)


# Hàm lấy giờ chấm công ra
def get_clock_out(current, on_duty, minute_must_ci):
    global list_datetime
    res = ""
    flag = False
    # Lấy thời gian on duty, off duty
    hour_on_duty = int(on_duty.split(":")[0])
    minute_on_duty = int(on_duty.split(":")[1])
    # Lấy thời gian bắt đầu và kết thúc
    start = current + timedelta(hours=hour_on_duty, minutes=minute_on_duty)
    end = current + timedelta(hours=hour_on_duty, minutes=minute_on_duty) + timedelta(days=1) - timedelta(
        minutes=minute_must_ci, seconds=1)
    # Lặp và lấy đúng thời gian chấm công ra, không có trả về None
    for i, current in enumerate(list_datetime):
        if datetime_in_range(start, end, current):
            res = current
            flag = True
            continue
        if flag:
            list_datetime = list_datetime[i:]
            break
    return str(res)


# Hàm xuất file excel
def output_excel(input_detail, path):
    # Xác định số hàng và cột lớn nhất trong file excel cần tạo
    row = len(input_detail)
    column = len(input_detail[0])

    # Tạo một workbook mới và active nó
    wb = openpyxl.Workbook()
    ws = wb.active

    # Dùng vòng lặp for để ghi nội dung từ input_detail vào file Excel
    for i in range(0, row):
        for j in range(0, column):
            try:
                v = input_detail[i][j]
                ws.cell(column=j + 1, row=i + 1, value=v)
            except:
                break

    # Lưu lại file Excel
    wb.save(path)


# Hàm main
def main():
    global list_datetime

    # Mở file config
    f = open("config.json", encoding="utf-8")

    # Trả về dạng dictionary
    config = json.load(f)

    # Lặp và kết nối các máy chấm công từ file config
    attendances = []
    conn = None

    for x in config["time_attendance"]:
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
    list_timetable = []

    for shift in config["shift"]:
        list_on_duty.append(shift["work_start"])
        list_off_duty.append(shift["work_end"])
        list_timetable.append(shift["shift_code"])

    # Đóng file config
    f.close()

    # Tiêu đề excel
    header = [["AC-No.", "No.", "Name", "Date", "Timetable", "On duty", "Off duty", "Clock In", "Clock Out", "Normal",
               "Real time", "Late", "Early", "Absent", "OT Time", "Work Time", "Exception", "Must C/In", "Must C/Out",
               "Department", "NDays", "WeekEnd", "Holiday", "ATT_Time", "NDays_OT", "WeekEnd_OT", "Holiday_OT"]]
    excel = header

    # Đưa vào khoảng thời gian muốn lấy ngày chấm công
    date_from = datetime.datetime(2023, 2, 27)
    date_to = datetime.datetime(2023, 3, 1)
    date = get_date(date_from, date_to)

    # Lặp lấy thông tin và đưa vào mảng
    for i in range(len(users)):
        ac_no = users[i].user_id
        no = users[i].user_id
        name = users[i].name
        list_datetime_ = get_attendances_of_user(ac_no, attendances, date_from, date_to)

        for j in range(len(date)):
            current = date[j]

            for k in range(len(list_on_duty)):
                row = []
                list_datetime = list_datetime_
                on_duty = list_on_duty[k]
                off_duty = list_off_duty[k]
                time_table = list_timetable[k]
                clock_in = get_clock_in(current, on_duty, off_duty, minute_must_ci)
                clock_out = get_clock_out(current, on_duty, minute_must_ci) if clock_in is not None else None

                row.append(ac_no)
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
    path = 'DuLieuChamCong.xlsx'

    # Xuất file
    output_excel(excel, path)
    print("Xuất file Excel thành công")


main()
