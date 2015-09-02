__author__ = 'zloy'

import xlrd
import config
from os import listdir,path

def read_day(sheet,start,num_of_groups,day=1):
    """
    :param sheet: передается страница
    :param start: передается левый верхний угол ячейки дня
    :param day: если первый день в неделе - передается двойка
    :return: вощзращается список структур из данных ячеек
    """

    def check_merge(sheet,row,column):
        merges=sheet.merged_cells
        for carret in merges:
            if carret[0]<=row<=carret[1]:
                if carret[2]<=column<=carret[3]:
                    return True
                    break
        else:
            return False

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

def load_kurses(dir,sht):

    select_curse_size = lambda name: config.kurses_size[name.split('_')[1]]
    lst=listdir(dir)
    weeks=[]
    for i in lst:
        if path.splitext(i)[1]=='.xls':
            weeks.append(i)
    massive=[]
    dates=[]
    for kurse in weeks:
        addr=dir+kurse
        book = xlrd.open_workbook(addr, encoding_override="cp1252", formatting_info=True)
        sheet = book.sheet_by_index(sht)
        tmp=load_week(sheet,select_curse_size(kurse))
        massive.append(tmp)
        for i in range(7,25):
            if sheet.cell_value(i,1)!='':
                dates.append(sheet.cell_value(i,1))
        for i in range(27,43):
            if sheet.cell_value(i,1)!='':
                dates.append(sheet.cell_value(i,1))
        for i in range(46,62):
            if sheet.cell_value(i,1)!='':
                dates.append(sheet.cell_value(i,1))
        for i in range(65,81):
            if sheet.cell_value(i,1)!='':
                dates.append(sheet.cell_value(i,1))
        for i in range(84,100):
            if sheet.cell_value(i,1)!='':
                dates.append(sheet.cell_value(i,1))
        for i in range(103,119):
            if sheet.cell_value(i,1)!='':
                dates.append(sheet.cell_value(i,1))
        for i in range(len(dates)):
            tmp=dates[i]
            for k in range(10):
                tmp=tmp.replace('  ',' ')
            tmp=tmp.strip()
            dates[i]=tmp
    dates = list(set(dates))
    dates.sort()
    return (massive,dates)

def parse_name_string(st):
    names=st.split(':')
    names[1]=names[1].split('|')
    names[1]=tuple(names[1])
    return names

def load_names():
    names = []
    for i in config.names.items():
        names.append(i)
    names.sort()
    return names