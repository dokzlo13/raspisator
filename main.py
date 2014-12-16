__author__ = 'zloy'
__version__='0.1a'

import xlrd

import inpt,out,analytics

import argparse


try:

    parser = argparse.ArgumentParser(
        prog='raspisator',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=('''\
            Генератор расписания для преподавателей
            ---------------------------------------
            '''),
         epilog=('''\

            Если не указан один из параметров, он будет запрошен в интерактивном режиме

            Программа имеет два необходимых параметра:
                -i Каталог с файлами расписаний
                -n Файл с именами преподавателей

            Файлы расписаний должны быть представлены в формате: <something>_<номер_курса>_<something>.xls

            Файл с именами должен иметь следующий синтаксис:
                <Имя, отображаемое в таблице>:<имена>|<в>|<расписании>
            При указании параметра -m фамилии преподавателей водятся вручную

        '''))

    parser.add_argument('-v', '--version', action='version', version='%(prog)s ' + __version__)
    parser.add_argument('-i','--inpt', action="store", help="Каталог с расписаниями")
    parser.add_argument('-o','--out', action="store", help="Файл для сохранения")
    parser.add_argument('-m','--manual', action="store_true", help="Ручной режим ввода фамилий")
    parser.add_argument('-n','--names', action="store", help="Файл списка имен")
    parser.add_argument('-s','--sheet', action="store", type=int, help="Страница, на которой находится желаемая неделя расписания")
    parser.add_argument('-c','--color', action="store_true", help="Раскраска цветом")
    args = parser.parse_args()

    if not args.inpt:
        args.inpt=input('Введите адрес каталога с расписаниями: ')
    if not args.out:
        args.out='./individual.xlsx'
    if not args.sheet:
        args.sheet=int(input('Введите номер страницы: '))
    if not args.manual:
        if not args.names:
            args.names=input('Введите файл с именами для поиска (manual - ручной режим) : ')
            if args.names=='manual':
                args.manual=True
                names=[]
    else:
        names=[]
    if not args.color:
        args.color=False

    print('Создание структуры расписаний')
    massive,dates=inpt.load_kurses(args.inpt,args.sheet) #Загрузка расписаний и дат из таблиц
    print('Загрузка имен')

    if not args.manual:
        names=inpt.load_names(args.names) #возвращает [['Имя',('имя')],...]
    else:
        st='1'
        while st!='NULL':
            st=input('Введите строку для поиска в формате <Имя, отображаемое в таблице>:<имена>|<в>|<расписании>, чтобы закончить введите NULL:\n')
            if st!='NULL':
                names.append(inpt.parse_name_string(st))

    print('Всего преподавателей', len(names))
    print('Запись каркаса расписания')
    book=out.write_base(args.out,dates,len(names)) #Запись основы в таблицу

    clr=2
    for human in range(len(names)): #цикл по количеству преподавателей
        book=out.write_name(book,names[human][0],human) #запись имени преподавателя
        for i in range(6):#цикл дней для каждого препода
            tmp=analytics.find_prep(massive,names[human][1],i) #поиск пар препода в указанный день
            if args.color:
                book=out.write_data(book,tmp,i,human,clr) #запись в таблицу этого дня
            else:
                book=out.write_data(book,tmp,i,human) #запись в таблицу этого дня

        print('Для преподавателя {0} записано расписание'.format(names[human][0]))

        if clr==12:
            clr=2
        else:
            clr+=1

    out.save(book)#сохранение книги

except KeyboardInterrupt:
    print('\nАварийное завершение!\n')