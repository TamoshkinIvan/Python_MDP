# Python_MDP
authors: Tamoshkin Ivan and Mizev Artem 
### О программе 
Программа предназначенная для расчетов МДП и АДП согласно Правилам определения максимально допустимых и аварийно допустимых перетоков активной мощности в контролируемых сечениях. 
### Необходимые требования 
* Python (x32) 3.7 (или выше)
* Pywin32 установленный в виртуальное окружение 
#### Необходимые файлы:
| Param | Description |
| ------ | ------ |
| -rg2 | Путь к файлу режима |
| -rg2template | Путь к файлу шаблона режима |
| -bg | Путь к json файлу с сечениями |
| -outages | Путь к json файлу с возмущениями |
| -pfvv | Путь к csv файлу с нормативными возмущениями |
### Пример использования
```sh
MaxPowerFlow -rg2 "Tests\assets\regime.rg2" -pfvv "Tests\assets\vector.csv" -rg2template "src\assets\rastr_templates\режим.rg2" -bg "Tests\assets\flowgate.json" -outages "Tests\assets\faults.json" 
```
#### Результаты:
```sh
• 20% Pmax запас в нормальном режиме:     2217
• 15% Ucr запас в нормальном режиме:      2778
• ДДТН в нормальном режиме:               1708
• 8% Pmax запас в послеаварийном режиме:  2131
• 10% Ucr запас в послеаварийном режиме:  2318
• АДТН в послеаварийном режиме:	          1474
```
