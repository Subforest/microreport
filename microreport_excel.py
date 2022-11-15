import datetime as d
import os
from tkinter.filedialog import Open

import openpyxl
import psycopg2

probeg = '6003'
"""Пробег, км"""
potreb = '6003'
"""Потребление электроэнергии , тысяч кВт*час"""
motor_komp = '0000'
"""Кол-во циклов включения-выключения мотор-компрессора, раз"""
dveri = '6005'
"""Кол-во циклов открытия-закрытия дверей, раз"""

path_out = 'out' 
"""Путь к выходным файлам"""

def datetime_to_int(dt:d.date) -> int:
    """ Формирует measuredatetime для запросов к базе (количество секунд с 1970 до начала дня)
    dt: date"""

    return (round((d.datetime.combine(dt,d.time(0,0,0)) - d.datetime(1970,1,1)).total_seconds()*1000))



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
    ss = ';Дата;'+d.date.today().strftime("%d.%m.%Y")+';'
    # file_out.write(ss)
    ss = 'Статистика по вагонам за день:;;'
    # file_out.write(ss)

    # Получение данных за день
    curs.execute("select wagonid,msgid,value from wagondata w where wagonid <> '30000' and msgid in ('6003','6005') and "  
                    "measuredatetime >= 1666742400000 and measuredatetime <2500286400000 order by wagonid ,msgid ,measuredatetime")
    num_wag=''
    row_wag = {}
    for row in curs:
        if num_wag!=row[0]:
            num_wag = row[0]
            row_wag[num_wag] = {'num':num_wag,probeg:0,dveri:[0,0,0,0,0,0,0,0]}
            # print(row_wag)
        if row[1]==probeg:
            row_wag[num_wag][probeg] = row_wag[num_wag][probeg] + int(row[2])
        elif row[1]==dveri: 
            # Обработка Кол-во циклов открытия-закрытия дверей
            r = row[2].split()
            for i in range(8):
                row_wag[num_wag][dveri][i] = row_wag[num_wag][dveri][i] + int(r[i])

    # for row in row_wag:
    #     print(row_wag[row])
        
    wb = openpyxl.load_workbook('template_microreport.xlsx')

    print(wb.sheetnames)

    sheet = wb['За месяц']
    print(sheet['C2'].value)

    sheet['C2'].value = '00001'

    f_out_name = 'Новокосино_'+d.date.today().strftime("%Y_%m_%d")+'.xlsx'
    if not(os.path.exists(path_out)):
        os.makedirs(path_out)
    # file_out = open(file=path_out+'/'+f_out_name,mode='w')
    wb.save(path_out+'/'+f_out_name)
    # ss = 'Номер вагона:;00000;'

    server.stop()
