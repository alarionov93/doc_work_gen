import xlrd, xlwt
import datetime as dt
import sys, traceback

dates = [6,10,15,17,22,25,30]
themes = [
	'Afdsf',
	'gfrf',
	'asfasf',
	'sggsarg',
	'sfsdf',
	'asfsf',
	'asrfs'
]

wb = xlwt.Workbook()
sht = wb.add_sheet('Пасищаимасдь')

for r_id in range(len(dates)):
	try:
		sht.write(r_id, 0, '%s.%s.%s' % (2020, 11, dates[r_id]))
		sht.write(r_id,1,themes[r_id])
		sht.write(r_id,2,'2')
		sht.write(r_id,3,'Ларионов А.А.')
	except Exception:
		print(traceback.print_exc(file=sys.stdout))
wb.save('test_12.xlsx')