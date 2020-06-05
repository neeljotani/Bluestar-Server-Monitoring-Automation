import os
import shutil
from time import gmtime,strftime
from pytz import timezone
from datetime import datetime
import time




India = timezone('Asia/Kolkata')
ind_time = datetime.now(India)
date1 = ind_time.strftime("%d_%m_%Y")
hour = int(ind_time.strftime("%I"))
print(hour)
print(hour == 10)
src_dir="Report Format/Format.xlsx"
dst_dir="Bluesr_New_Prod_Server_hourly_report_"+date1+".xlsx"
if hour == 10:
	shutil.copy(src_dir,dst_dir)
else:
	print("Not the first report")
