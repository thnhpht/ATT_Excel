import json
from zk import ZK, const
import zk
import unidecode

# Mở file config
f = open("config.json", encoding="utf-8")
  
# Trả về dạng dictionary
config = json.load(f)

conn = None
zk = ZK('192.168.1.201', port=4370)
try:
    conn = zk.connect()
    for user in config["users"]:
        conn.set_user(user_id=user["id"], name=unidecode.unidecode(user["name"]), privilege=const.USER_DEFAULT)
except Exception as e:
    print ("Process terminate : {}".format(e))
finally:
    if conn:
        conn.disconnect()

# Đóng file config
f.close()