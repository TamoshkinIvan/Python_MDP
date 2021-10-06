# Отключение необходимых линий
def LineOFF(rastr, row):
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