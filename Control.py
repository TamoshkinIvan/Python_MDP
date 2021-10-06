# Обнуление режима, выставление необходимого контроля по I или V, перенос
# данных по току
def Control(rastr,ShablonRegime, criteria, AV=False):
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