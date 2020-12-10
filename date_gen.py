import datetime as dt
import sys
import xlrd, xlwt
import traceback

target_dates = []
try:
	wd_1 = int(sys.argv[1])
	wd_2 = int(sys.argv[2])
	if len(sys.argv) > 3 or wd_1 not in range(0,6) or wd_2 not in range(0,6):
		print("Usage: date_gen.py weekday_1 weekday_2")
		exit(1)
	target_weekdays = (wd_1, wd_2)
except ValueError:
	print("Weekdays use digits from 0 to 6 only")
	exit(2)
except IndexError:
	print("Wrong number of parameters, trying 1 weekday")
	target_weekdays = [wd_1]

now_month = dt.datetime.now().month
now_year = dt.datetime.now().year

try:
	dt.date(now_year, now_month, 31)
	max_month_cnt = 31
except ValueError:
	try:
		dt.date(now_year, now_month, 30)
		max_month_cnt = 30
	except ValueError:
		try:
			dt.date(now_year, now_month, 29)
			max_month_cnt = 29
		except ValueError:
			max_month_cnt = 28
for d in range(1, max_month_cnt+1):
	target_dates.append(dt.date(now_year, now_month, d))

print(max_month_cnt)

res_dates = []
for d in target_dates:
	if d.weekday() in target_weekdays:
		res_dates.append("%d.%d.%d" % (d.day, d.month, d.year))

# open('test.csv', 'w').writelines(["%s,\n" % x for x in res_dates])
print(res_dates)
print("Total dates: %s" % len(res_dates))

