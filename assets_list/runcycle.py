import os
import sys
import time
from shareholders_info import  MyRequests

root_dir=os.path.dirname(os.path.abspath(__file__))
flag_run=os.path.join(root_dir,"run")
flag_first=os.path.join(root_dir,"first")
run_py=os.path.join(root_dir,"shareholders_info.py")
print("flag_run={}".format(flag_run))
print("flag_first={}".format(flag_first))
print("run_py={}".format(run_py))

while True:
    if os.path.exists(flag_first):
        os.remove(flag_first)
        mr = MyRequests()
        mr.process_run()
    else:
        if os.path.exists(flag_run):
            os.remove(flag_run)
            time.sleep(900)
            mr = MyRequests()
            mr.process_run()


