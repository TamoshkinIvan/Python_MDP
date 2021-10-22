# Отключение необходимых линий
from pandas import DataFrame
from win32com.client import Dispatch
from typing import Optional


def line_off(rastr: Dispatch, row: DataFrame):
    """
    Функция используемая для отключения линий
    rastr - рассчитываемый режим
    row - DatаFrame с отключаемой линией
    """
    vetv = rastr.Tables('vetv')
    sta = vetv.Cols('sta')
    ip = vetv.Cols('ip')
    iq = vetv.Cols('iq')
    np = vetv.Cols('np')
    # Формируем послеаварийную схему
    for i in range(vetv.Size):
        if ip.Z(i) == row['ip'] and iq.Z(
                i) == row['iq'] and np.Z(i) == row['np']:
            sta.SetZ(i, row['sta'])


# Сбор данных по перетокам в сечении
def get_power_flow(sechen: Dispatch) -> float:
    """
    функция определяет переток по сечению
    return
    mdp - предельный переток
    """
    mdp = 0
    for i in range(sechen.Size):
        mdp += abs(sechen.Cols('pl').Z(i))
    return mdp

# Утяжеление до конца, вычисление МДП по критерию
def calculation_mdp(rastr: Dispatch, fluctuations:float, sechen: Dispatch,
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
    mdp = get_power_flow(sechen)
    if contingency is not None:
        tpf = get_power_flow(sechen)
        mdp *= k_zap
        toggle = rastr.GetToggle()
        j = 1
        while tpf > mdp:
            toggle.MoveOnPosition(len(toggle.GetPositions()) - j)
            tpf = get_power_flow(sechen)
            j += 1
        vetv = rastr.Tables('vetv')
        vetv.SetSel('ip={_ip}&iq={_iq}&np={_np}'.format(_ip=contingency['ip'],
                                                        _iq=contingency['iq'],
                                                        _np=contingency['np']))
        vetv.Cols('sta').Calc(0)
        rastr.rgm('p')
        tpf = get_power_flow(sechen)
        return round(tpf - fluctuations)
    return round(mdp * k_zap - fluctuations)