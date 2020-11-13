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
now_month = dt.datetime.now().year

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
for d in range(1, max_month_cnt):
	target_dates.append(dt.date(now_year, now_month, d))

res_dates = []
for d in target_dates:
	if d.weekday() in target_weekdays:
		res_dates.append("%d.%d.%d" % (d.day, d.month, d.year))

open('test.csv', 'w').writelines(["%s,\n" % x for x in res_dates])

#sheet = xlrd.open_workbook('test.xlsx')
#if __name__ == '__main__':
#	print('hello')
# dates = [6,10,15,17,22,25,30]
themes = [
	'Введение в понятие алгоритма. Виды алгоритмов',
	'Обзор актуальных языков программирования',
	'Операторы ввода/вывода и ариф. действия',
	'Условные операторы',
	'Создание текстовой игры с нелинейным сюжетом',
	'Операторы повторения',
	'Создание матрицы',
	'Массивы',
	'Введение в предмет. Схематехника',
	'Радиокомпоненты и принципиальные схемы',
	'Составление схемы звукового генератора',

]

wb = xlwt.Workbook()
sht = wb.add_sheet('1 Учёт пройденного материала')

for r_id in range(len(res_dates)):
	try:
		sht.write(r_id, 0, res_dates[r_id])
		sht.write(r_id, 1, themes[r_id])
		sht.write(r_id, 2, '2')
		sht.write(r_id, 3, 'Ларионов А.А.')
	except Exception:
		print(traceback.print_exc(file=sys.stdout))

sht = wb.add_sheet('2 Учёт посещаемости')
# TODO: import from csv?

wb.save('test_12.xlsx')





