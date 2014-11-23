__author__ = 'zloy'

# -*- coding: UTF-8 -*-
import xlsxwriter

def write_base(dates):

    zero=0

    workbook = xlsxwriter.Workbook('individual.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})

    worksheet.set_column(0,0,20)
    worksheet.set_column(zero+1,30,5)
    worksheet.set_row(zero,20)

    worksheet.merge_range(zero,0,zero+1,0, 'Ф.И.О', bold)

    pos=1
    worksheet.merge_range(zero,pos+0,zero,pos+3+1, dates[0])
    worksheet.merge_range(zero,pos+4+1,zero,pos+7+2, dates[1])
    worksheet.merge_range(zero,pos+4+4+2,zero,pos+7+4+3, dates[2])
    worksheet.merge_range(zero,pos+4+4+4+3,zero,pos+7+4+4+4, dates[3])
    worksheet.merge_range(zero,pos+4+4+4+4+4,zero,pos+7+4+4+4+5, dates[4])
    worksheet.merge_range(zero,pos+4+4+4+4+4+5,zero,pos+7+4+4+4+4+6, dates[5])

    par=('1-2','3-4','5-6','7-8','9-10')
    k=0
    for i in range(1,31):
        worksheet.write(zero+1,i,par[k])
        if k==4:
            k=0
        else:
            k+=1

    return workbook


def write_name(workbook,name,man):
    worksheet=workbook.worksheets()[0]
    pos_row = man*4+2
    worksheet.write(pos_row,0,name)
    return workbook

def write_data(workbook,paras,day=0,man=0,color_set=0):
    worksheet=workbook.worksheets()[0]

    pos_st = day*5+1
    pos_row = man*4+2

    #Установка цвета для дня
    colors=['#0000FF','#00FFFF','#800000','#008000','#00FF00','#FF00FF','#000080','#FFFF00','#800080','#00FF00','#FF6600','#FF0000']
    color = workbook.add_format()
    color.set_pattern(1)
    if color_set==0 or color_set+1>len(colors):
        color.set_bg_color('#FFFFFF')
    else:
        color.set_bg_color(colors[color_set-1])
    '''
    if day==0:
        pos=1
    elif day==1:
        pos=6
    elif day==2:
        pos=11
    elif day==3:
        pos=16
    elif day==4:
        pos=21
    elif day==5:
        pos=26
    '''
    for i in range(5):
        for j in range(len(paras)):
            if paras[j][0]==i+1:
                i=pos_st+i
                worksheet.write(pos_row,i,paras[j][1],color)
                worksheet.write(pos_row+2,i,paras[j][2],color)
                worksheet.write(pos_row+3,i,paras[j][4].replace('.0',''),color)
                worksheet.write(pos_row+1,i,paras[j][5].replace('.0',''),color)

    return workbook

def save(workbook):
    workbook.close()