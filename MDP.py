# Импортируем необходимые библеотеки
import win32com.client
import pandas as pd
import matplotlib as mp
import numpy as np
import sys

ShablonRegime = 'Shablons/режим.rg2'
ShablonTracktoria = 'Shablons/траектория утяжеления.ut2'
ShablonSechenia = 'Shablons/сечения.sch'

fluctuations = 30


# Утяжеление до конца, вычисление МДП по критерию
def CalculationMDP(Kzap):
    rastr.rgm('p')
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')
    MDP = 0
    sechen = rastr.Tables('grline')
    for i in range(sechen.Size):
        MDP += abs(pl.Z(i))
    return round(MDP * Kzap - fluctuations)


# Обнуление режима, выставление необходимого контроля по I или V, перенос
# данных по току
def Control(criteria, AV=False):
    rastr.Load(1, 'Rejime/regime.rg2', ShablonRegime)
    # Увеличим количество итераций
    ut_common = rastr.Tables('ut_common')
    ut_common.Cols('iter').SetZ(0, 200)
    # Подготовимся к включению контроля по току и напряжению при утяжелении
    enable_contr = ut_common.Cols('enable_contr')
    dis_i_contr = ut_common.Cols('dis_i_contr')
    dis_p_contr = ut_common.Cols('dis_p_contr')
    dis_v_contr = ut_common.Cols('dis_v_contr')
    if (criteria == 'V'):
        enable_contr_set = enable_contr.SetZ(0, 1)
        # Отключим контроль I
        dis_i_contr_set = dis_i_contr.SetZ(0, 1)
        # Отключим контроль P
        dis_p_contr_set = dis_p_contr.SetZ(0, 1)
        # Включим контроль V
        dis_v_contr_set = dis_v_contr.SetZ(0, 0)
    elif (criteria == 'I'):
        enable_contr_set = enable_contr.SetZ(0, 1)
        dis_p_contr_set = dis_p_contr.SetZ(0, 1)
        dis_v_contr_set = dis_v_contr.SetZ(0, 1)
        dis_i_contr_set = dis_i_contr.SetZ(0, 0)
        vetv = rastr.Tables('vetv')
        i_dop_ob = vetv.Cols('i_dop')
        i_dop_r = vetv.Cols('i_dop_r')
        if AV:
            i_dop_ob = vetv.Cols('i_dop_ob')
            i_dop_r = vetv.Cols('i_dop_r_av')
        contr_i = vetv.Cols('contr_i')
        # Неизвестно почему, но данные по АДП находятся в расчетной чатсти,
        # перенесем их в столбец с ДДТН_доп
        for i in range(vetv.Size):
            i_dop_ob.SetZ(i, i_dop_r.Z(i))
            if i_dop_ob.Z(i) != 0:
                contr_i.SetZ(i, 1)


# Отключение необходимых линий
def LineOFF(row):
    vetv = rastr.Tables('vetv')
    sta = vetv.Cols('sta')
    ip = vetv.Cols('ip')
    iq = vetv.Cols('iq')
    np = vetv.Cols('np')
    # Формируем послеаварийную схему
    for i in range(vetv.Size):
        if ip.Z(i) == row['ip'] and iq.Z(
                i) == row['iq'] and np.Z(i) == row['np']:
            sta.SetZ(i, 1)


rastr = win32com.client.Dispatch("Astra.Rastr")

# Загрузим файсл с режимом
rastr.Load(1, 'Rejime/regime.rg2', ShablonRegime)
# Загрузим файсл с траекторией
rastr.Save('Rejime/траектория утяжеления.ut2', ShablonTracktoria)
rastr.Load(1, 'Rejime/траектория утяжеления.ut2', ShablonTracktoria)
# Загрузим файсл с сечением
rastr.Save('Rejime/сечения.sch', ShablonSechenia)
rastr.Load(1, 'Rejime/сечения.sch', ShablonSechenia)

# Прочитаем файлы возмущений, сечения и траектории
faults = pd.read_json('Rejime/faults.json')
flowgate = pd.read_json('Rejime/flowgate.json')
vector = pd.read_csv('Rejime/vector.csv')

faults = faults.T
flowgate = flowgate.T

LoadTrajectory = vector[vector['variable'] == 'pn']
LoadTrajectory = LoadTrajectory.rename(
    columns={
        'variable': 'pn',
        'value': 'pn_value',
        'tg': 'pn_tg'})
GenTrajectory = vector[vector['variable'] == 'pg']
GenTrajectory = GenTrajectory.rename(
    columns={
        'variable': 'pg',
        'value': 'pg_value',
        'tg': 'pg_tg'})

vector = pd.merge(left=GenTrajectory, right=LoadTrajectory,
                  left_on='node', right_on='node', how='outer').fillna(0)

# Таблица траектории утяжеления
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
rastr.Save('Rejime/траектория утяжеления1.ut2', ShablonTracktoria)

# Таблица сечений
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
rastr.Save('Rejime/сечения.sch', ShablonSechenia)

# Обеспечение нормативного коэффициента запаса статической апериодической
# устойчивости по активной мощности в контролируемом сечении в нормальной
# схеме.
Control('P')
MDP_1 = CalculationMDP(0.8)
print(MDP_1)

# Обеспечение нормативного коэффициента запаса
# статической устойчивости по напряжению в узлах нагрузки в нормальной схеме.
Control('V')
MDP_2 = CalculationMDP(1)
print(MDP_2)

# Обеспечение нормативного коэффициента запаса
# статической апериодической устойчивости
# по активной мощности в контролируемом сечении в
# послеаварийных режимах после нормативных возмущений.
for index, row in faults.iterrows():
    Control('P')
    # Отключим линию
    LineOFF(row)
    # Определим значение перетока
    MDP_3 = CalculationMDP(0.92)
    print(MDP_3)

# Обеспечение нормативного коэффициента запаса статической
# устойчивости по напряжению в узлах нагрузки в послеаварийных режимах
# после нормативных возмущений.
# Итерируемся по строкам в датафрейме с нормативными возмущениями
for index, row in faults.iterrows():
    Control('V')
    LineOFF(row)
    # Определим значение перетока
    MDP_4 = CalculationMDP(1)
    print(MDP_4)

# Токое в норм схеме
# Определим значение перетока
Control('I')
MDP_5_1 = CalculationMDP(1)
print(MDP_5_1)

# Токое в ПАр
# Определим значение перетока
for index, row in faults.iterrows():
    Control('I', True)
    LineOFF(row)
    # Определим значение перетока
    MDP_5_2 = CalculationMDP(1)
    print(MDP_5_2)
