
# VPC конвектор

Позволяет обрабатывать файлы форматы .vpc используемые в модуле PCS версии OML до 3.0


## Установка

1. Клонируйте проект **vpcconv** c GitHub

2. Создайте виртуальное окружение

* Windows
```bash
python -m venv venv
```

* Mac OS и Linux
```bash
python3 -m venv venv
```

3. Активируйте виртуальное окружение
 * Windows
```bash
source venv/Scripts/activate
```

* Mac OS и Linux
```bash
source venv/bin/activate
```

4. Установите необходимые библиотеки
```bash
pip install -r requirements.txt
```



## Подготовка файлы формата .vpc

1. Измените имя и формата файла на ```input.txt```.
2. Удалите лишние строки из файла до первой строки с текстос /ITEM.
```     
        /ITEM
         /UNAME   =U105ZI02904B                     'PRES CTRL   in 105-V-200'
         /UTYPE   =PVI
         /VARS
             /VAR
                 /VNAME   =PV                               IN
                 /SCALE   =1                0                '%'
                 /VTYPE   =2        0        0
                 /ATTR
                 /REFERS
                     /REFER   =1        203      MODEL3      U105PV02904B      pos             0        0
```
3. Поместите в директорию с программой ```vpcconv.py``` и запустите её.
4. В результате выполнения программы создадуться два файла: .csv и .xlsx.