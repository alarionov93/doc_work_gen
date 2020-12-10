import datetime as dt
import sys
import xlrd, xlwt
import traceback


teacher_name = 'Ларионов А.А.'
full_teacher_name = 'Ларионов Александр Андреевич'
#group_name = 'Б-2'

def set_col_width(sheet, col_idx, width):
	try:
		sheet.col(col_idx).width = width
	except Exception:
		print(traceback.print_exc(file=sys.stdout))	


def get_csv(url):
	urld = 'https://docs.google.com/spreadsheets/d/1sb209ihX0VhHRCYvBQjc4T8NuDsHHv8kBVmdmGJORmI/export?exportFormat=csv&gid=0'
	# TODO: requests needed!

target_dates = []
try:
	wd_1 = int(sys.argv[1])
	wd_2 = int(sys.argv[2])
	if len(sys.argv) > 6 or wd_1 not in range(0,6) or wd_2 not in range(0,6):
		print("Usage: date_gen.py weekday_1 weekday_2 filename_in_csv out_filename")
		exit(1)
	target_weekdays = (wd_1, wd_2)
except ValueError:
	print("Weekdays use digits from 0 to 6 only")
	exit(2)
except IndexError:
	print("Wrong number of parameters, trying 1 weekday")
	target_weekdays = [wd_1]

try:
	filename = sys.argv[3]
except IndexError:
	print("No filename of csv file")
	exit(3)

try:
	out_filename = sys.argv[4]
except IndexError:
	print("No out filename")
	exit(4)

try:
	group_name = sys.argv[5]
except IndexError:
	print("No group name passed")
	exit(5)


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

# print(max_month_cnt)

res_dates = []
celebration_days = []
# celebration_days.append(dt.date(2020,11,4)) # day of constitution!!!
if 1: # TODO: use argv from command line!
	for cd in range(20,max_month_cnt+1):
		celebration_days.append(dt.date(2020, 12, cd))
for d in target_dates:
	if d.weekday() in target_weekdays and d not in celebration_days:
		# res_dates.append("%d.%d.%d" % (d.day, d.month, d.year))
		res_dates.append(d)

# print(res_dates)

# open('test.csv', 'w').writelines(["%s,\n" % x for x in res_dates])

#sheet = xlrd.open_workbook('test.xlsx')

# TODO: MAIN TODO!!! откуда то брать смещение в темах из учебной программы!!!
# И автоматизировать загрузку из гугл документов этих сраных!!!
#
# itp2 - 8
# itb2,6 - 21
# itpr - 9
# 6 занятий сгенерировано на 7.12.20
n = 8 # the last string number not for use
themes_all = [x.strip('\n') for x in open('themes.%s.csv' % filename, 'r').readlines()]
themes = themes_all[n:]
print('Themes shift (new n value): %s' % (len(themes) + n )) # TODO: check this value!!
# exit(0)

wb = xlwt.Workbook()
sht = wb.add_sheet('1 Учёт пройденного материала')

date_style = xlwt.XFStyle()
date_style.num_format_str = 'DD.MM.YYYY'

# print(res_dates)

for r_id in range(len(res_dates)):
	try:
		sht.write(r_id, 0, res_dates[r_id], date_style)
		sht.write(r_id, 1, themes[r_id])
		max_wdth = sht.col(1).width
		if len(themes[r_id]*367) > max_wdth:
			set_col_width(sht, 1, len(themes[r_id]*267))
		sht.write(r_id, 2, 2)
		sht.write(r_id, 3, teacher_name)
		sht.col(3).width = len(teacher_name*267)
	except Exception:
		print(traceback.print_exc(file=sys.stdout))

sht = wb.add_sheet('2 Учёт посещаемости')

# TODO: import from csv? Yes!
p = open('%s - Sheet1.csv' % filename).read()

# merge some cells
max_len_of_merged = 37
start_col = 2
start_row = 3

try:
	sht.write_merge(1, 3, 0, 1, '№ п/п Фамилия, имя учащегося')
	sht.write_merge(1, 1, 2, max_len_of_merged, 'IT, IT%s, %s' % (group_name, full_teacher_name))
except Exception:
	print(traceback.print_exc(file=sys.stdout))	

# dates line
for c_id in range(len(res_dates)):
	try:
		sht.write(start_row, c_id + start_col, res_dates[c_id].day)
	except Exception:
		print(traceback.print_exc(file=sys.stdout))

# 37 - len of merged top cells with group and teacher name
for c_id in range(max_len_of_merged):
	sht.col(c_id + start_col).width = 800


idx = start_row + 1
st_id = 1
for student in p.split("\n")[1:]:
	vals = student.split(',')
	st_name = student.split(',')[0]
	p = vals[1:]
	
	sht.write(idx, 0, st_id)
	set_col_width(sht, 0, 900)
	sht.write(idx, 1, st_name)
	max_wdth = sht.col(1).width
	if len(st_name*267) > max_wdth:
		sht.col(1).width = len(st_name*367)

	for c_id in range(len(res_dates)):
		try:
			if p[c_id] == 'x':
				val = ' '
			elif p[c_id] == 'n':
				val = 'н'
			else:
				val = 'н'
			sht.write(idx, c_id + 2, val)
		except Exception:
			print(traceback.print_exc(file=sys.stdout))

	idx += 1
	st_id += 1
# for r_id in range(len(res_dates)):
# 	try:
# 		sht.write()
# 	except Exception:
# 		print(traceback.print_exc(file=sys.stdout))


wb.save('%s.xlsx' % out_filename)





