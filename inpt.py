__author__ = 'zloy'

#to fix#
# в ячейках пар для нескольких потоков считывается только одна группа

import xlrd

def load_dates(addr,sht): #убрать пробелы
    book = xlrd.open_workbook(addr, encoding_override="cp1252")
    sheet = book.sheet_by_index(sht)
    days=[]
    pos=27

    #print(sheet.cell_value(7,1))
    #print(sheet.cell_value(27,1))
    #print(sheet.cell_value(46,1))
    #print(sheet.cell_value(65,1))
    #print(sheet.cell_value(84,1))
    #print(sheet.cell_value(103,1))

    days.append(sheet.cell_value(7,1))
    for i in range(5):
        days.append(sheet.cell_value(pos,1))
        pos+=19
    return days


def check_merge(sheet,row,column):
    merges=sheet.merged_cells
    for carret in merges:
        if carret[0]<=row<=carret[1]:
            if carret[2]<=column<=carret[3]:
                return True
                break
    else:
        return False


def read_day(sheet,start,num_of_groups,day=1):
    """
    :param sheet: передается страница
    :param start: передается левый верхний угол ячейки дня
    :param day: если первый день в неделе - передается двойка
    :return: вощзращается список структур из данных ячеек
    """
    rasp=[]
    group=(start[0]-day)
    numr_pary=1

    for i in range(4):
        for j in range(num_of_groups):
            para=(numr_pary,(sheet.cell_value(start[0]+i,start[1]+j)), str(sheet.cell_value(start[0]+i+1,start[1]+j)), str(sheet.cell_value(start[0]+i+2,start[1]+j)),str(sheet.cell_value(start[0]+i+3,start[1]+j)).replace('.0',''),str(sheet.cell_value(group,start[1]+j)).replace('.0',''))
            if not para[1]==para[2]==para[3]==para[4]=='':
                rasp.append(para)

            else:
                if check_merge(sheet,start[0]+i,start[1]+j):
                    para=(numr_pary,(sheet.cell_value(start[0]+i,start[1]+j-1)), str(sheet.cell_value(start[0]+i+1,start[1]+j-1)), str(sheet.cell_value(start[0]+i+2,start[1]+j-1)),str(sheet.cell_value(start[0]+i+3,start[1]+j-1)).replace('.0',''),str(sheet.cell_value(group,start[1]+j)).replace('.0',''))
                    if not para[1]==para[2]==para[3]==para[4]=='':
                        if numr_pary!=4: #[ИНДУС] если 4 пара, то в потоке она не ведется, выяснить почему merge барахлит
                            rasp.append(para)

        start=(start[0]+3,start[1])
        numr_pary+=1
    return rasp


def load_week(sheet,num_of_groups):
    days=[]
    days.append(read_day(sheet,(9,2),num_of_groups,2))
    pos=28
    for i in range(5):
        days.append(read_day(sheet,(pos,2),num_of_groups))
        pos+=19
    return(days)

def load_kurses(dir,sht=12):
    from os import listdir,path
    #dir='./docs1/'
    lst=listdir(dir)
    weeks=[]
    for i in lst:
        if path.splitext(i)[1]=='.xls':
            weeks.append(i)
    massive=[]
    for i in weeks:
        addr=dir+i
        book = xlrd.open_workbook(addr, encoding_override="cp1252", formatting_info=True)
        sheet = book.sheet_by_index(sht)
        if i.split('_')[1]=='1':
            tmp=load_week(sheet,19)
        if i.split('_')[1]=='2':
            tmp=load_week(sheet,13)
        if i.split('_')[1]=='3':
            tmp=load_week(sheet,11)
        if i.split('_')[1]=='4':
            tmp=load_week(sheet,10)
        if i.split('_')[1]=='5':
            tmp=load_week(sheet,6)
        if i.split('_')[1]=='mag':
            tmp=load_week(sheet,11)

        massive.append(tmp)
    dates=load_dates(dir+weeks[0],sht)
    return (massive,dates)


def parse_name_string(st):
    names=st.split(':')
    names[1]=names[1].split('|')
    names[1]=tuple(names[1])
    return names

def load_names(addr):
    """
    Формат конфига имен:
    Имя_в_таблице:Имена, В, Расписании
    двоеточие не отделяется пробелами
    Каждое из имен распиания отделяется запятой и пробелом (, )

    :param addr: адрес конфига имен
    :return: возвращается список в формате [[Имя_в_таблице, (Имена, в, расписании)], [...], ...]
    """
    try:
        file=open(addr)
        raw=file.read()
        file.close()
    except FileNotFoundError:
        print('Файл не найден!')
        exit(-1)

    lines=raw.split('\n')

    '''Реализовать комментарии конфига имен
    for j in range(len(lines)):
        if lines[j][0]=='#':
            lines.pop(j)
    '''
    names=[]
    for j in lines:
        names.append(parse_name_string(j))

    return names
