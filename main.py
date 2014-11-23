__author__ = 'zloy'

##################TODO LIST:#####################
##(???)Пофиксить считывание пар на несколько групп DONE
##добавить пятый курс для расписания DONE
##Отладить список фамилий IN PROGRESS
##Сравнить получаемый план с настоящим DONE
##Пофиксить считывание дат из расписания
##Пофиксить комментарии конфига имен
##Реализовать обработку коммандной строки
#################################################

import inpt,out,analytics, xlrd

massive,dates=inpt.load_kurses() #Загрузка расписаний и дат из таблиц
book=out.write_base(dates) #Запись основы в таблицу
names=inpt.load_names('./txt') #возвращает [['Имя',('имя')],...]

clr=2

for human in range(len(names)): #цикл по количеству преподавателей
    book=out.write_name(book,names[human][0],human) #запись имени преподавателя
    for i in range(6):#цикл дней для каждого препода
        tmp=analytics.find_prep(massive,names[human][1],i) #поиск пар препода в указанный день
        book=out.write_data(book,tmp,i,human) #запись в таблицу этого дня

    if clr==12:
        clr=2
    else:
        clr+=1

out.save(book)#сохранение книги
