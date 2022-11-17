import datetime as d
import os
import argparse
from tkinter.filedialog import Open

import openpyxl
import psycopg2

mes_spis = ['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь']

probeg = '6003'
"""Пробег, км"""
potreb = '0001'
"""Потребление электроэнергии , тысяч кВт*час"""
motor_komp = '0000'
"""Кол-во циклов включения-выключения мотор-компрессора, раз"""
dveri = '6005'
"""Кол-во циклов открытия-закрытия дверей, раз"""
buksy = '0010'
"""Подробный статус неисправностей букс"""

path_out = 'out' 
"""Путь к выходным файлам"""
name_line = 'Новокосино'
"""Наименование линии метрополитена"""

one_day = d.timedelta(days=1)
"""Один день для добавления в цикле"""

parse = argparse.ArgumentParser()
parse.add_argument("-m",action="store_true",dest="m"
                    ,help="Отчёт за предыдущий месяц")
parse.add_argument("-w",action="store_true"
                    ,help="Отчёт за предыдущую неделю")
args = parse.parse_args()

def datetime_to_int(dt:d.date) -> int:
    """ Формирует measuredatetime для запросов к базе (количество секунд с 1970 до начала дня)
    dt: date"""

    return (round((d.datetime.combine(dt,d.time(0,0,0)) - d.datetime(1970,1,1)).total_seconds()*1000))

if not(os.path.exists(path_out)):
    os.makedirs(path_out)


# Для туннелирования. для сервера - убрать
from sshtunnel import SSHTunnelForwarder

with SSHTunnelForwarder(
     ('admin.mip.tools', 30022),
     ssh_private_key="/home/apodlesskih/.ssh/ssh_mip_ns_user",
     ### in my case, I used a password instead of a private key
     ssh_username="ns-user",
    #  ssh_password="<mypasswd>",
     remote_bind_address=('localhost', 5432)) as server:
    server.start()
    print ("server connected")
# Конец туннелирования 
    
    params = {
         'database': 'traindb',
         'user': 'trainuser',
         'password': 'train',
         'host': 'localhost',
         'port': server.local_bind_port
         }
    #  params = settings.params_postgres
    conn = psycopg2.connect(**params)
    curs = conn.cursor()
    print ("database connected")

    date_now = d.date.today()
    if args.w:
        print("неделя")
        week_day = date_now.weekday()
        date_end = date_now - d.timedelta(days=week_day+1)
        date_begin = date_end - d.timedelta(days=7)
        f_out_name = name_line+'_'+date_begin.strftime("%Y_%m_%d")+"-"+date_end.strftime("%Y_%m_%d")+'.xlsx'
        title_period = 'Суммарные показатели за неделю с '+date_begin.strftime("%Y.%m.%d")+' по '+date_end.strftime("%Y_%m_%d")

        date_end = date_end + one_day # Поправили дату чтобы использовать условие <
        row_wag = {} # За период
        row_wag_d ={} # По дням

        # Перебираем по одному дню, начиная с даты начала
        while date_begin<date_end:
            q = ("select wagonid,msgid,value from wagondata w where wagonid <> '30000' and msgid in ('6003','6005') and "  
                "measuredatetime >= "+str(datetime_to_int(date_begin))+" and measuredatetime <"+str(datetime_to_int(date_begin+one_day))
                +" order by wagonid ,msgid ,measuredatetime")
            curs.execute(q)
            num_wag=''
            for row in curs:
                if num_wag!=row[0]:
                    num_wag = row[0]
                    if num_wag not in row_wag.keys():
                        row_wag[num_wag] = {'num':num_wag,probeg:0,potreb:0,motor_komp:0,dveri:[0,0,0,0,0,0,0,0]}
                if row[1]==probeg:
                    row_wag[num_wag][probeg] = row_wag[num_wag][probeg] + int(row[2])
                elif row[1]==potreb:
                    row_wag[num_wag][potreb] = row_wag[num_wag][potreb] + int(row[2])
                elif row[1]==motor_komp:
                    row_wag[num_wag][motor_komp] = row_wag[num_wag][motor_komp] + int(row[2])
                elif row[1]==dveri: 
                    # Обработка Кол-во циклов открытия-закрытия дверей
                    r = row[2].split()
                    for i in range(8):
                        row_wag[num_wag][dveri][i] = row_wag[num_wag][dveri][i] + int(r[i])
            date_begin = date_begin+one_day    
        #Вывод периода в файл
        num_str = 7
        wb = openpyxl.load_workbook('template_microreport.xlsx')
        sheet = wb['За период']
        sheet['C2'].value = title_period

        for row in row_wag:
            sheet['C'+str(num_str)].value = row
            sheet['D'+str(num_str)].value = str(row_wag[row][probeg])
            sheet['E'+str(num_str)].value = str(row_wag[row][potreb])
            sheet['F'+str(num_str)].value = str(row_wag[row][motor_komp])
            sheet['G'+str(num_str)].value = str(row_wag[row][dveri][0])
            sheet['H'+str(num_str)].value = str(row_wag[row][dveri][1])
            sheet['I'+str(num_str)].value = str(row_wag[row][dveri][2])
            sheet['J'+str(num_str)].value = str(row_wag[row][dveri][3])
            sheet['K'+str(num_str)].value = str(row_wag[row][dveri][4])
            sheet['L'+str(num_str)].value = str(row_wag[row][dveri][5])
            sheet['M'+str(num_str)].value = str(row_wag[row][dveri][6])
            sheet['N'+str(num_str)].value = str(row_wag[row][dveri][7])
            num_str+=1
        wb.save(path_out+'/'+f_out_name)
        # print(date_begin)
        # print(date_end)

    # for row in row_wag:
    #     print(row_wag[row])


    server.stop()
