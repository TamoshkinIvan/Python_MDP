# Импортируем необходимые библеотеки
import win32com.client
import pandas as pd
from pandas import DataFrame
import control
import line_off
from typing import Optional
import csv
import time
start_time = time.time()
shablon_regime = 'Shablons/режим.rg2'
shablon_tracktoria = 'Shablons/траектория утяжеления.ut2'
shablon_sechenia = 'Shablons/сечения.sch'
fluctuations = 30


# Сбор данных по перетокам в сечении
def get_power_flow() -> float:
    """
    функция определяет переток по сечению
    """
    mdp = 0
    for i in range(sechen.Size):
        mdp += abs(sechen.Cols('pl').Z(i))
    return mdp


# Утяжеление до конца, вычисление МДП по критерию
def calculation_mdp(
        k_zap: float, contingency: Optional[DataFrame] = None) -> float:
    """
    Функция расчитывает предельный переток
    k_zap - коэффициент запаса
    contingency - отключаемый элемент
    return - возвращает предельный переток по заданному критерию
    """
    rastr.rgm('p')
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')
    mdp = get_power_flow()
    if contingency is not None:
        tpf = get_power_flow()
        mdp *= k_zap
        toggle = rastr.GetToggle()
        j = 1
        while tpf > mdp:
            toggle.MoveOnPosition(len(toggle.GetPositions()) - j)
            tpf = get_power_flow()
            j += 1
        vetv = rastr.Tables('vetv')
        vetv.SetSel('ip={_ip}&iq={_iq}&np={_np}'.format(_ip=contingency['ip'],
                                                        _iq=contingency['iq'],
                                                        _np=contingency['np']))
        vetv.Cols('sta').Calc(0)
        rastr.rgm('p')
        tpf = get_power_flow()
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
faults = pd.read_json('regime/faults.json').T
flowgate = pd.read_json('regime/flowgate.json').T
vector = pd.read_csv('regime/vector.csv')


def csv_to_dict(path: str) -> [dict]:
    """ Функция производит парсинг сsv в словарь"""
    dict_list = []
    with open(path, newline='') as csv_data:
        csv_dic = csv.DictReader(csv_data)
        # Creating empty list and adding dictionaries (rows)
        for row in csv_dic:
            dict_list.append(row)
    return dict_list


def add_node_tr(node_num: int, recalc_tan: int) -> int:
    """ Функция функция добавляет в таблицу траектрии узлы
    и устанавливает tg
    node_num - номер узла
    recalc_tan - учет тангенса tg
    """
    i = rastr.Tables('ut_node').size
    rastr.Tables('ut_node').AddRow()
    rastr.Tables('ut_node').Cols('ny').SetZ(i, node_num)
    rastr.Tables('ut_node').Cols('tg').SetZ(i, recalc_tan)
    return i


def set_node_tr_param(node_id: int,
                      param: str,
                      value: float) -> None:
    """ Функция функция добавляет в таблицу траектрии парамемты утяжеления"""
    rastr.Tables('ut_node').Cols(param).SetZ(node_id, value)


reader = csv_to_dict('regime/vector.csv')
node_id_map = {}
for row in reader:
    node = row.get('node', 0)
    if node not in node_id_map:
        node_id = add_node_tr(node, row.get('tg', 0))
        node_id_map[node] = node_id
    else:
        node_id = node_id_map[node]
    variable = row.get('variable', 'pn')
    set_node_tr_param(
        node_id, variable, float(row.get('value', 0)))

# Таблица сечений
sechen = rastr.Tables('grline')
i = 0
ns_init = 1

for index, row in flowgate.iterrows():
    sechen.AddRow()
    sechen.Cols('ns').SetZ(i, ns_init)
    sechen.Cols('ip').SetZ(i, row['ip'])
    sechen.Cols('iq').SetZ(i, row['iq'])
    i += 1
rastr.Save('regime/сечения.sch', shablon_sechenia)

# Обеспечение нормативного коэффициента запаса статической апериодической
# устойчивости по активной мощности в контролируемом сечении в нормальной
# схеме.
control.control(rastr, shablon_regime, 'P')
mdp_1 = calculation_mdp(0.8)
print("20% Pmax запас в нормальном режиме: " + str(mdp_1))

# Обеспечение нормативного коэффициента запаса
# статической устойчивости по напряжению в узлах нагрузки в нормальной схеме.
control.control(rastr, shablon_regime, 'V')
mdp_2 = calculation_mdp(1)
print("15% Ucr запас в нормальном режиме: " + str(mdp_2))

# Обеспечение нормативного коэффициента запаса
# статической апериодической устойчивости
# по активной мощности в контролируемом сечении в
# послеаварийных режимах после нормативных возмущений.
mdp_3 = []
for index, contingency in faults.iterrows():
    control.control(rastr, shablon_regime, 'P')
    # Отключим линию
    line_off.line_off(rastr, contingency)
    # Определим значение перетока
    mdp_3.append(calculation_mdp(0.92, contingency))
print("8% Pmax запас в послеаварийном режиме: " + str(min(mdp_3)))

# Обеспечение нормативного коэффициента запаса статической
# устойчивости по напряжению в узлах нагрузки в послеаварийных режимах
# после нормативных возмущений.
# Итерируемся по строкам в датафрейме с нормативными возмущениями
mdp_4 = []
for index, contingency in faults.iterrows():
    control.control(rastr, shablon_regime, 'V')
    line_off.line_off(rastr, contingency)
    # Определим значение перетока
    mdp_4.append(calculation_mdp(1, contingency))
print("10% Ucr запас в послеаварийном режиме: " + str(min(mdp_4)))

# Токое в норм схеме
# Определим значение перетока
control.control(rastr, shablon_regime, 'I')
mdp_5_1 = calculation_mdp(1)
print("ДДТН в нормальном режиме: " + str(mdp_5_1))

# Токое в ПАр
# Определим значение перетока
mdp_5_2 = []
for index, contingency in faults.iterrows():
    control.control(rastr, shablon_regime, 'I', True)
    line_off.line_off(rastr, contingency)
    # Определим значение перетока
    mdp_5_2.append(calculation_mdp(1, contingency))
print("АДТН в послеаварийном режиме: " + str(min(mdp_5_2)))
print("--- %s seconds ---" % (time.time() - start_time))