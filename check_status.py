from zk import ZK, const
import zk

conn = None
zk = ZK('192.168.1.201', port=4370)
try:
    conn = zk.connect()
except Exception as e:
    print("Ngưng hoạt động")
finally:
    if conn:
        print("Đang hoạt động")
        conn.disconnect()
        