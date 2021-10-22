import win32com.client
from win32com.client import Dispatch
import csv


# Обнуление режима, выставление необходимого контроля по I или V, перенос
# данных по току
def control(rastr: Dispatch, shablon_regime: Dispatch,
            criteria: str, post_fault_mode=False) -> None:
    """
    Функция для выставления контролируемого параметра
    в процессе утяжеления
    rastr - рассчитываемый режим
    shablon_regime - шаблон режима
    criteria - критерий расчета
    post_fault_mode - индикатор расчета для ПАР
    """
    rastr.Load(1, 'regime/regime.rg2', shablon_regime)
    # Увеличим количество итераций
    ut_common = rastr.Tables('ut_common')
    ut_common.Cols('iter').SetZ(0, 200)
    # Подготовимся к включению контроля по току и напряжению при утяжелении
    enable_contr = ut_common.Cols('enable_contr')
    dis_i_contr = ut_common.Cols('dis_i_contr')
    dis_p_contr = ut_common.Cols('dis_p_contr')
    dis_v_contr = ut_common.Cols('dis_v_contr')
    enable_contr_set = enable_contr.SetZ(0, 1)
    # Отключим контроль P
    dis_p_contr_set = dis_p_contr.SetZ(0, 1)
    if criteria == 'V':
        # Отключим контроль I
        dis_i_contr_set = dis_i_contr.SetZ(0, 1)
        # Включим контроль V
        dis_v_contr_set = dis_v_contr.SetZ(0, 0)
    elif criteria == 'I':
        dis_v_contr_set = dis_v_contr.SetZ(0, 1)
        dis_i_contr_set = dis_i_contr.SetZ(0, 0)
        vetv = rastr.Tables('vetv')
        i_dop_ob = vetv.Cols('i_dop')
        i_dop_r = vetv.Cols('i_dop_r')
        if post_fault_mode:
            i_dop_ob = vetv.Cols('i_dop_ob')
            i_dop_r = vetv.Cols('i_dop_r_av')
        contr_i = vetv.Cols('contr_i')
        # Неизвестно почему, но данные по АДП находятся в расчетной чатсти,
        # перенесем их в столбец с ДДТН_доп
        for i in range(vetv.Size):
            i_dop_ob.SetZ(i, i_dop_r.Z(i))
            if i_dop_ob.Z(i) != 0:
                contr_i.SetZ(i, 1)


def csv_to_dict(path: str) -> [dict]:
    """ Функция производит парсинг сsv в словарь
        path - пусть к файлу с траектрией утяжеления
        return
        dict_list - траекторию утяжеления
    """
    dict_list = []
    with open(path, newline='') as csv_data:
        csv_dic = csv.DictReader(csv_data)
        # Creating empty list and adding dictionaries (rows)
        for row in csv_dic:
            dict_list.append(row)
    return dict_list


def add_node_tr(rastr: Dispatch, node_num: int, recalc_tan: int) -> int:
    """ Функция функция добавляет в таблицу траектрии узлы
    и устанавливает tg
    node_num - номер узла
    recalc_tan - учет тангенса tg
    return - i - номер строки в таблице утяжеления
    """
    i = rastr.Tables('ut_node').size
    rastr.Tables('ut_node').AddRow()
    rastr.Tables('ut_node').Cols('ny').SetZ(i, node_num)
    rastr.Tables('ut_node').Cols('tg').SetZ(i, recalc_tan)
    return i


def set_node_tr_param(rastr: Dispatch,
                      node_id: int,
                      param: str,
                      value: float) -> None:
    """ Функция функция добавляет в таблицу траектрии параметры утяжеления
        node_id - параметр узла
        param - Параметр утяжеления pg/pn
        value - Приращение pg/pn
        return - None
    """
    rastr.Tables('ut_node').Cols(param).SetZ(node_id, value)
