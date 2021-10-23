# Импортируем необходимые библеотеки
import win32com.client
import pandas as pd
import time
import preparation
import calculation

if __name__ == '__main__':
    start_time = time.time()
    shablon_regime = 'Shablons/режим.rg2'
    shablon_tracktoria = 'Shablons/траектория утяжеления.ut2'
    shablon_sechenia = 'Shablons/сечения.sch'
    fluctuations = 30

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

    reader = preparation.csv_to_dict('regime/vector.csv')
    node_id_map = {}
    for row in reader:
        node = row.get('node', 0)
        if node not in node_id_map:
            node_id = preparation.add_node_tr(rastr, node, row.get('tg', 0))
            node_id_map[node] = node_id
        else:
            node_id = node_id_map[node]
        variable = row.get('variable', 'pn')
        preparation.set_node_tr_param(rastr,
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
    preparation.control(rastr, shablon_regime, 'P')
    mdp_1 = calculation.calculation_mdp(rastr, fluctuations, sechen, 0.8)
    print("20% Pmax запас в нормальном режиме: " + str(mdp_1))

    # Обеспечение нормативного коэффициента запаса
    # статической устойчивости по напряжению в узлах нагрузки в нормальной схеме.
    preparation.control(rastr, shablon_regime, 'V')
    mdp_2 = calculation.calculation_mdp(rastr, fluctuations, sechen, 1)
    print("15% Ucr запас в нормальном режиме: " + str(mdp_2))

    # Обеспечение нормативного коэффициента запаса
    # статической апериодической устойчивости
    # по активной мощности в контролируемом сечении в
    # послеаварийных режимах после нормативных возмущений.
    mdp_3 = []
    for index, contingency in faults.iterrows():
        preparation.control(rastr, shablon_regime, 'P')
        # Отключим линию
        calculation.line_off(rastr, contingency)
        # Определим значение перетока
        mdp_3.append(calculation.calculation_mdp(rastr, fluctuations, sechen, 0.92, contingency))
    print("8% Pmax запас в послеаварийном режиме: " + str(min(mdp_3)))

    # Обеспечение нормативного коэффициента запаса статической
    # устойчивости по напряжению в узлах нагрузки в послеаварийных режимах
    # после нормативных возмущений.
    # Итерируемся по строкам в датафрейме с нормативными возмущениями
    mdp_4 = []
    for index, contingency in faults.iterrows():
        preparation.control(rastr, shablon_regime, 'V', True)
        calculation.line_off(rastr, contingency)
        # Определим значение перетока
        mdp_4.append(calculation.calculation_mdp(rastr, fluctuations, sechen, 1, contingency))
    print("10% Ucr запас в послеаварийном режиме: " + str(min(mdp_4)))

    # Токое в норм схеме
    # Определим значение перетока
    preparation.control(rastr, shablon_regime, 'I')
    mdp_5_1 = calculation.calculation_mdp(rastr, fluctuations, sechen, 1)
    print("ДДТН в нормальном режиме: " + str(mdp_5_1))

    # Токое в ПАр
    # Определим значение перетока
    mdp_5_2 = []
    for index, contingency in faults.iterrows():
        preparation.control(rastr, shablon_regime, 'I', True)
        calculation.line_off(rastr, contingency)
        # Определим значение перетока
        mdp_5_2.append(calculation.calculation_mdp(rastr, fluctuations, sechen, 1, contingency))
    print("АДТН в послеаварийном режиме: " + str(min(mdp_5_2)))
    print("--- %s seconds ---" % (time.time() - start_time))
