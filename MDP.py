#Импортируем необходимые библеотеки
import win32com.client
import pandas as pd
import sys

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog

import form

class ExampleApp(QtWidgets.QMainWindow, form.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.connectActions()

if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()


rastr = win32com.client.Dispatch("Astra.Rastr")

# Загрузим файсл с режимом
rastr.Load(1, 'regime.rg2', 'G:/rastrwin3/RastrWin3/SHABLON/режим.rg2')
# Загрузим файсл с траекторией
rastr.Save('траектория утяжеления.ut2', 'G:/rastrwin3/RastrWin3/SHABLON/траектория утяжеления.ut2')

rastr.Load(1, 'траектория утяжеления.ut2', 'G:/rastrwin3/RastrWin3/SHABLON/траектория утяжеления.ut2')
# Загрузим файсл с сечением
rastr.Save('сечения.sch', 'G:/rastrwin3/RastrWin3/SHABLON/сечения.sch')
rastr.Load(1, 'сечения.sch', 'G:/rastrwin3/RastrWin3/SHABLON/сечения.sch')
#Значение RG_KOD 1 соотвествует режиму "Загрузить"
# Проверим файл
result = rastr.rgm('p')
#Вывод 0, следовательно расчет завершился успешно
print(result)

#Прочитаем файлы возмущений, сечения и траектории
faults = pd.read_json('faults.json')
flowgate = pd.read_json('flowgate.json')
vector = pd.read_csv('vector.csv')





flowgate = flowgate.T
print(flowgate)


LoadTrajectory = vector[vector['variable'] == 'pn']
LoadTrajectory = LoadTrajectory.rename(columns = {'variable':'pn', 'value':'pn_value', 'tg':'pn_tg'})
GenTrajectory = vector[vector['variable'] == 'pg']
GenTrajectory = GenTrajectory.rename(columns = {'variable':'pg', 'value':'pg_value', 'tg':'pg_tg'})


vector = pd.merge(left = GenTrajectory, right = LoadTrajectory,
                              left_on = 'node', right_on = 'node', how = 'outer').fillna(0)


#Таблица траектории утяжеления
ut_node = rastr.Tables('ut_node')
tip = ut_node.Cols('tip')
ny = ut_node.Cols('ny')
pn = ut_node.Cols('pn')
pg = ut_node.Cols('pg')
tg = ut_node.Cols('tg')
i = 0
for index, row in vector.iterrows():
    rastr.Tables('ut_node').AddRow()
    rastr.Tables('ut_node').Cols('ny').SetZ(i, row['node'])
    if pd.notnull(row['pg']):
        rastr.Tables('ut_node').Cols('pg').SetZ(i, row['pg_value'])
        rastr.Tables('ut_node').Cols('tg').SetZ(i, row['pg_tg'])
    if pd.notnull(row['pn']):
        rastr.Tables('ut_node').Cols('pn').SetZ(i, row['pn_value'])
        rastr.Tables('ut_node').Cols('tg').SetZ(i, row['pn_tg'])
    i = i + 1
rastr.Save('траектория утяжеления1.ut2', 'G:/rastrwin3/RastrWin3/SHABLON/траектория утяжеления.ut2')
ut_node.Size


#Таблица сечений
sechen = rastr.Tables('grline')
ns = sechen.Cols('ns')
ip = sechen.Cols('ip')
iq = sechen.Cols('iq')
pl = sechen.Cols('pl')
i = 0
ns_init = 1

for index, row in flowgate.iterrows():
    sechen.AddRow()
    ns.SetZ(i, ns_init)
    ip.SetZ(i, row['ip'])
    iq.SetZ(i, row['iq'])
    i += 1
rastr.Save('сечения.sch', 'G:/rastrwin3/RastrWin3/SHABLON/сечения.sch')
(sechen.Size)

#Обеспечение нормативного коэффициента запаса статической апериодической
#устойчивости по активной мощности в контролируемом сечении в нормальной схеме.

rastr.Load(1, 'regime.rg2', 'G:/rastrwin3/RastrWin3/SHABLON/режим.rg2')
ut_common = rastr.Tables('ut_common')

I_max = ut_common.Cols('iter')
I_max.SetZ(0, 200)

if rastr.ut_utr('i') > 0:
    rastr.ut_utr('')
print(rastr.rgm('p'))
#Определим  значение перетока
i = 0
MDP = 0
while i < sechen.Size:
    MDP +=  abs(pl.Z(i))
    i += 1
MDP_1 = round(MDP * 0.8  - 30)
print (MDP_1)

