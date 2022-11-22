import datetime as d
import os
import argparse

import openpyxl
from openpyxl.styles import Border,Alignment, Side, Font
import psycopg2

mes_spis = ['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь']
probeg = '6008'
"""Пробег, км"""
potreb = '6009'
"""Потребление электроэнергии , тысяч кВт*час"""
motor_komp = '6004'
"""Кол-во циклов включения-выключения мотор-компрессора, раз"""
dveri = '6005'
"""Кол-во циклов открытия-закрытия дверей, раз"""
buksy = '0010'
"""Подробный статус неисправностей букс"""

sql_head = "select wagonid,msgid,value from wagondata w where wagonid <> '30000' and msgid in ('"+probeg+"','"+potreb+"','"+motor_komp+"','"+dveri+"','"+buksy+"') and "
"""Начало SQL запроса с основными условиями"""
sql_tail = " order by wagonid ,msgid ,measuredatetime"
"""Конец SQL запроса с постобработкой"""

path_out = 'out' 
"""Путь к выходным файлам"""
name_line = 'Новогиреево'
"""Наименование линии метрополитена"""

one_day = d.timedelta(days=1)
"""Один день для добавления в цикле"""

parse = argparse.ArgumentParser()
parse.add_argument("-m",type=int,default=0,dest="m"
                    ,help="Отчёт за предыдущий месяц")
parse.add_argument("-w",type=int,default=0
                    ,help="Отчёт за предыдущую неделю")
parse.add_argument("-d",type=int,default=0
                    ,help="Отчёт за сегодня")
args = parse.parse_args()

def datetime_to_int(dt:d.date) -> int:
    """ Формирует measuredatetime для запросов к базе (количество милисекунд с 1970 до начала дня)
    dt: date"""

    return (round((d.datetime.combine(dt,d.time(0,0,0)) - d.datetime(1970,1,1)).total_seconds()*1000))

def datetime_to_int_str(dt:d.date) -> str:
    """ Формирует measuredatetime для запросов к базе (количество милисекунд с 1970 до начала дня)
    dt: date"""

    return (str(round((d.datetime.combine(dt,d.time(0,0,0)) - d.datetime(1970,1,1)).total_seconds()*1000)))

def xlsx_report(curs,date_begin:d.date,date_end:d.date,title_period:str,f_out_name:str):
    """ Создаёт и заполняет файл xlsx на заданный период"""
    date_end = date_end + one_day # Поправили дату чтобы использовать условие <
    row_wag = {} # За период
    row_wag_d ={} # По дням

    # Перебираем по одному дню, начиная с даты начала
    while date_begin<date_end:
        q = (sql_head  
            +"measuredatetime >= "+datetime_to_int_str(date_begin)
            +" and measuredatetime <"+datetime_to_int_str(date_begin+one_day)
            +sql_tail)
        curs.execute(q)
        num_wag=''
        date_b_str=date_begin.strftime("%Y.%m.%d")
        for row in curs:
            if num_wag!=row[0]:
                num_wag = row[0]
                if num_wag not in row_wag.keys():
                    row_wag[num_wag] = {'num':num_wag,probeg:0,
                                       potreb:0,
                                   motor_komp:0,
                                        dveri:[0,0,0,0,0,0,0,0],
                                        buksy:[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]}
                    row_wag_d[num_wag] = {date_b_str:{probeg:0,
                                                      potreb:0,
                                                   motor_komp:0,
                                                        dveri:[0,0,0,0,0,0,0,0],
                                                        buksy:[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]}
                                        }
            if date_b_str not in row_wag_d[num_wag].keys():
                row_wag_d[num_wag][date_b_str] = {probeg:0,
                                                  potreb:0,
                                              motor_komp:0,
                                                   dveri:[0,0,0,0,0,0,0,0],
                                                   buksy:[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]}

            if row[1]==probeg:
                row_wag[num_wag][probeg] += float(row[2])
                row_wag_d[num_wag][date_b_str][probeg] += float(row[2])
            elif row[1]==potreb:
                row_wag[num_wag][potreb] += float(row[2])
                row_wag_d[num_wag][date_b_str][potreb] += float(row[2])
            elif row[1]==motor_komp:
                row_wag[num_wag][motor_komp] += int(row[2])
                row_wag_d[num_wag][date_b_str][motor_komp] += int(row[2])
            elif row[1]==dveri: 
                # Обработка Кол-во циклов открытия-закрытия дверей
                r = row[2].split()
                for i in range(8):
                    row_wag[num_wag][dveri][i] += int(r[i])
                    row_wag_d[num_wag][date_b_str][dveri][i] += int(r[i])
            elif row[1]==buksy:
                # Обработка Кол-во срабатываний 
                r = row[2].split()
                for i in range(24):
                    row_wag[num_wag][buksy][i] += int(r[i])
                    row_wag_d[num_wag][date_b_str][buksy][i] += int(r[i])
        date_begin = date_begin+one_day    
    #Вывод периода в файл
    num = 6
    wb = openpyxl.load_workbook('template_microreport.xlsx')
    sheet = wb['За период']
    sheet['A1'].value = title_period

    thin_border = Border(left=Side(style='thin'),
                 right=Side(style='thin'),
                 top=Side(style='thin'),
                 bottom=Side(style='thin'))

    for row in sorted(row_wag.keys()):
        num_str = str(num)
        sheet['A'+num_str].value = row
        sheet['B'+num_str].value = round(row_wag[row][probeg],2)
        sheet['C'+num_str].value = round((row_wag[row][potreb]/1000),2)
        sheet['D'+num_str].value = row_wag[row][motor_komp]
        sheet['E'+num_str].value = row_wag[row][dveri][0]
        sheet['F'+num_str].value = row_wag[row][dveri][1]
        sheet['G'+num_str].value = row_wag[row][dveri][2]
        sheet['H'+num_str].value = row_wag[row][dveri][3]
        sheet['I'+num_str].value = row_wag[row][dveri][4]
        sheet['J'+num_str].value = row_wag[row][dveri][5]
        sheet['K'+num_str].value = row_wag[row][dveri][6]
        sheet['L'+num_str].value = row_wag[row][dveri][7]
        sheet['M'+num_str].value = row_wag[row][buksy][0]
        sheet['N'+num_str].value = row_wag[row][buksy][1]
        sheet['O'+num_str].value = row_wag[row][buksy][2]
        sheet['P'+num_str].value = row_wag[row][buksy][3]
        sheet['Q'+num_str].value = row_wag[row][buksy][4]
        sheet['R'+num_str].value = row_wag[row][buksy][5]
        sheet['S'+num_str].value = row_wag[row][buksy][6]
        sheet['T'+num_str].value = row_wag[row][buksy][7]
        sheet['U'+num_str].value = row_wag[row][buksy][8]
        sheet['V'+num_str].value = row_wag[row][buksy][9]
        sheet['W'+num_str].value = row_wag[row][buksy][10]
        sheet['X'+num_str].value = row_wag[row][buksy][11]
        sheet['Y'+num_str].value = row_wag[row][buksy][12]
        sheet['Z'+num_str].value = row_wag[row][buksy][13]
        sheet['AA'+num_str].value = row_wag[row][buksy][14]
        sheet['AB'+num_str].value = row_wag[row][buksy][15]
        sheet['AC'+num_str].value = row_wag[row][buksy][16]
        sheet['AD'+num_str].value = row_wag[row][buksy][17]
        sheet['AE'+num_str].value = row_wag[row][buksy][18]
        sheet['AF'+num_str].value = row_wag[row][buksy][19]
        sheet['AG'+num_str].value = row_wag[row][buksy][20]
        sheet['AH'+num_str].value = row_wag[row][buksy][21]
        sheet['AI'+num_str].value = row_wag[row][buksy][22]
        sheet['AJ'+num_str].value = row_wag[row][buksy][23]
        num+=1
    for _row in sheet['A6':'AJ'+str(num-1)]:
        for _cell in _row:
            _cell.border = thin_border
            _cell.alignment= Alignment(horizontal='center')

    sheet = wb['Ежедневно']
    num = 6

    for num_sost in sorted(row_wag_d.keys()):
        num_str = str(num)
        sheet.merge_cells('A'+num_str+':AJ'+num_str)
        sheet['A'+num_str].value = 'Вагон '+num_sost
        sheet['A'+num_str].font = Font(bold=True, size=12)
        sheet['A'+num_str].alignment= Alignment(horizontal='center')
        sheet['A'+num_str].border = thin_border
        num+=1
        for row in row_wag_d[num_sost]:
            num_str = str(num)
            sheet['A'+num_str].value = row
            sheet['B'+num_str].value = round(row_wag_d[num_sost][row][probeg],2)
            sheet['C'+num_str].value = round((row_wag_d[num_sost][row][potreb]/1000),2)
            sheet['D'+num_str].value = row_wag_d[num_sost][row][motor_komp]
            sheet['E'+num_str].value = row_wag_d[num_sost][row][dveri][0]
            sheet['F'+num_str].value = row_wag_d[num_sost][row][dveri][1]
            sheet['G'+num_str].value = row_wag_d[num_sost][row][dveri][2]
            sheet['H'+num_str].value = row_wag_d[num_sost][row][dveri][3]
            sheet['I'+num_str].value = row_wag_d[num_sost][row][dveri][4]
            sheet['J'+num_str].value = row_wag_d[num_sost][row][dveri][5]
            sheet['K'+num_str].value = row_wag_d[num_sost][row][dveri][6]
            sheet['L'+num_str].value = row_wag_d[num_sost][row][dveri][7]
            sheet['M'+num_str].value = row_wag_d[num_sost][row][buksy][0]
            sheet['N'+num_str].value = row_wag_d[num_sost][row][buksy][1]
            sheet['O'+num_str].value = row_wag_d[num_sost][row][buksy][2]
            sheet['P'+num_str].value = row_wag_d[num_sost][row][buksy][3]
            sheet['Q'+num_str].value = row_wag_d[num_sost][row][buksy][4]
            sheet['R'+num_str].value = row_wag_d[num_sost][row][buksy][5]
            sheet['S'+num_str].value = row_wag_d[num_sost][row][buksy][6]
            sheet['T'+num_str].value = row_wag_d[num_sost][row][buksy][7]
            sheet['U'+num_str].value = row_wag_d[num_sost][row][buksy][8]
            sheet['V'+num_str].value = row_wag_d[num_sost][row][buksy][9]
            sheet['W'+num_str].value = row_wag_d[num_sost][row][buksy][10]
            sheet['X'+num_str].value = row_wag_d[num_sost][row][buksy][11]
            sheet['Y'+num_str].value = row_wag_d[num_sost][row][buksy][12]
            sheet['Z'+num_str].value = row_wag_d[num_sost][row][buksy][13]
            sheet['AA'+num_str].value = row_wag_d[num_sost][row][buksy][14]
            sheet['AB'+num_str].value = row_wag_d[num_sost][row][buksy][15]
            sheet['AC'+num_str].value = row_wag_d[num_sost][row][buksy][16]
            sheet['AD'+num_str].value = row_wag_d[num_sost][row][buksy][17]
            sheet['AE'+num_str].value = row_wag_d[num_sost][row][buksy][18]
            sheet['AF'+num_str].value = row_wag_d[num_sost][row][buksy][19]
            sheet['AG'+num_str].value = row_wag_d[num_sost][row][buksy][20]
            sheet['AH'+num_str].value = row_wag_d[num_sost][row][buksy][21]
            sheet['AI'+num_str].value = row_wag_d[num_sost][row][buksy][22]
            sheet['AJ'+num_str].value = row_wag_d[num_sost][row][buksy][23]
            for _row in sheet['A'+num_str:'AJ'+num_str]:
                for _cell in _row:
                    _cell.border = thin_border
                    _cell.alignment= Alignment(horizontal='center')
            num+=1




    wb.save(path_out+'/'+f_out_name)

if not(os.path.exists(path_out)):
    os.makedirs(path_out)
  
params = {
     'database': 'traindb',
     'user': 'trainuser',
     'password': 'train',
     'host': 'localhost',
     'port': 5432
     }
#  Для связи с базой на тестовом сервере нужно открыть ssh туннель     
#  ssh -L 5432:127.0.0.1:5432 ns-user@admin.mip.tools -p 30022
conn = psycopg2.connect(**params)
curs = conn.cursor()
print ("database connected")
date_now = d.date.today()
# -------------------------
date_begin = date_now
for i in range(args.d):
    date_begin = date_now - d.timedelta(days=i)
    date_end = date_begin + one_day
    f_out_name = name_line+'_За день_'+date_begin.strftime("%Y_%m_%d")+"-"+date_end.strftime("%Y_%m_%d")+'.xlsx'
    title_period = 'Суммарные показатели за день с '+date_begin.strftime("%Y.%m.%d")+' по '+date_end.strftime("%Y_%m_%d")
    xlsx_report(curs,date_begin,date_end,title_period,f_out_name)
#--------------------------
date_begin = date_now
for i in range(args.w):
    print("неделя")
    week_day = date_begin.weekday()
    date_end = date_begin - d.timedelta(days=week_day+1)
    date_begin = date_end - d.timedelta(days=6)
    f_out_name = name_line+'_'+date_begin.strftime("%Y_%m_%d")+"-"+date_end.strftime("%Y_%m_%d")+'.xlsx'
    title_period = 'Суммарные показатели за неделю с '+date_begin.strftime("%Y.%m.%d")+' по '+date_end.strftime("%Y_%m_%d")
    xlsx_report(curs,date_begin,date_end,title_period,f_out_name)
date_begin = date_now
for i in range(args.m):
    print('Месяц')
    date_end = d.date(date_begin.year,date_begin.month,1) - one_day
    date_begin = d.date(date_end.year,date_end.month,1)
    title_period = 'Суммарные показатели за '+mes_spis[date_begin.month-1]+' месяц'
    f_out_name = name_line+'_'+date_begin.strftime("%Y_%m")+'.xlsx'
    xlsx_report(curs,date_begin,date_end,title_period,f_out_name)
