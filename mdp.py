# Импортируем необходимые библеотеки
import win32com.client
import pandas as pd
from pandas import DataFrame
import control
import line_off

shablon_regime = 'Shablons/режим.rg2'
shablon_tracktoria = 'Shablons/траектория утяжеления.ut2'
shablon_sechenia = 'Shablons/сечения.sch'
fluctuations = 30


# Сбор данных по перетокам в сечении
def collector_mdp() -> float:
    """
    функция определяет переток по сечению
    """
    mdp = 0
    sechen = rastr.Tables('grline')
    for i in range(sechen.Size):
        mdp += abs(pl.Z(i))
    return mdp


# Утяжеление до конца, вычисление МДП по критерию
def calculation_mdp(k_zap: float, row: DataFrame, av: bool = False):
    """
    Функция расчитывает предельный переток
    k_zap - коэффициент запаса
    row - отключаемый элемент
    av - ПАР
    return - возвращает предельный переток по заданному критерию
    """
    rastr.rgm('p')
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')
    mdp = collector_mdp()
    if av:
        tpf = collector_mdp()
        mdp *= k_zap
        toggle = rastr.GetToggle()
        j = 1
        while tpf > mdp:
            toggle.MoveOnPosition(len(toggle.GetPositions()) - j)
            tpf = collector_mdp()
            j += 1
        vetv = rastr.Tables('vetv')
        vetv.SetSel('ip={_ip}&iq={_iq}&np={_np}'.format(_ip=row['ip'],
                                                        _iq=row['iq'],
                                                        _np=row['np']))
        vetv.Cols('sta').Calc(0)
        rastr.rgm('p')
        tpf = collector_mdp()
        return round(tpf - fluctuations)
    else:
        return round(mdp * k_zap - fluctuations)


rastr = win32com.client.Dispatch("Astra.Rastr")

# Загрузим файсл с режимом
rastr.Load(1, 'regime/regime.rg2', shablon_regime)
# Загрузим файсл с траекторией
rastr.Save('regime/траектория утяжеления.ut2', shablon_tracktoria)
rastr.Load(1, 'regime/траектория утяжеления.ut2', shablon_tracktoria)
# Загрузим файсл с сечением
rastr.Save('regime/сечения.sch', shablon_sechenia)
rastr.Load(1, 'regime/сечения.sch', shablon_sechenia)

# Прочитаем файлы возмущений, сечения и траектории
faults = pd.read_json('regime/faults.json')
flowgate = pd.read_json('regime/flowgate.json')
vector = pd.read_csv('regime/vector.csv')

faults = faults.T
flowgate = flowgate.T

load_trajectory = vector[vector['variable'] == 'pn']
load_trajectory = load_trajectory.rename(
    columns={
        'variable': 'pn',
        'value': 'pn_value',
        'tg': 'pn_tg'})
gen_trajectory = vector[vector['variable'] == 'pg']
gen_trajectory = gen_trajectory.rename(
    columns={
        'variable': 'pg',
        'value': 'pg_value',
        'tg': 'pg_tg'})

vector = pd.merge(left=gen_trajectory, right=load_trajectory,
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
rastr.Save('regime/траектория утяжеления1.ut2', shablon_tracktoria)

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
rastr.Save('regime/сечения.sch', shablon_sechenia)

# Обеспечение нормативного коэффициента запаса статической апериодической
# устойчивости по активной мощности в контролируемом сечении в нормальной
# схеме.
control.control(rastr, shablon_regime, 'P')
mdp_1 = calculation_mdp(0.8, None)
print("20% Pmax запас в нормальном режиме: " + str(mdp_1))

# Обеспечение нормативного коэффициента запаса
# статической устойчивости по напряжению в узлах нагрузки в нормальной схеме.
control.control(rastr, shablon_regime, 'V')
mdp_2 = calculation_mdp(1, None)
print("15% Ucr запас в нормальном режиме: " + str(mdp_2))

# Обеспечение нормативного коэффициента запаса
# статической апериодической устойчивости
# по активной мощности в контролируемом сечении в
# послеаварийных режимах после нормативных возмущений.
mdp_3 = []
for index, row in faults.iterrows():
    control.control(rastr, shablon_regime, 'P')
    # Отключим линию
    line_off.line_off(rastr, row)
    # Определим значение перетока
    mdp_3.append(calculation_mdp(0.92, row, True))
print("8% Pmax запас в послеаварийном режиме: " + str(min(mdp_3)))

# Обеспечение нормативного коэффициента запаса статической
# устойчивости по напряжению в узлах нагрузки в послеаварийных режимах
# после нормативных возмущений.
# Итерируемся по строкам в датафрейме с нормативными возмущениями
mdp_4 = []
for index, row in faults.iterrows():
    control.control(rastr, shablon_regime, 'V', True)
    line_off.line_off(rastr, row)
    # Определим значение перетока
    mdp_4.append(calculation_mdp(1, row, True))
print("10% Ucr запас в послеаварийном режиме: " + str(min(mdp_4)))

# Токое в норм схеме
# Определим значение перетока
control.control(rastr, shablon_regime, 'I')
mdp_5_1 = calculation_mdp(1, None)
print("ДДТН в нормальном режиме: " + str(mdp_5_1))

# Токое в ПАр
# Определим значение перетока
mdp_5_2 = []
for index, row in faults.iterrows():
    control.control(rastr, shablon_regime, 'I', True)
    line_off.line_off(rastr, row)
    # Определим значение перетока
    mdp_5_2.append(calculation_mdp(1, row, True))
print("АДТН в послеаварийном режиме: " + str(min(mdp_5_2)))
