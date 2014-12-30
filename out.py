__author__ = 'zloy'

# -*- coding: UTF-8 -*-
import xlsxwriter

def write_base(addr,dates,names=2):

    zero=0

    workbook = xlsxwriter.Workbook(addr)
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold': True})
    left_border=workbook.add_format({'left':5})
    right_border=workbook.add_format({'right':5})
    top_right_border=workbook.add_format({'top':1,'right':5})
    top_left_border=workbook.add_format({'top':1,'left':5})
    top_border=workbook.add_format({'top':1})

    worksheet.set_column(0,0,20)
    worksheet.set_column(zero+1,30,5)
    worksheet.set_row(zero,20)

    worksheet.merge_range(zero,0,zero+1,0, 'Ф.И.О', bold)

    pos=1

    worksheet.merge_range(zero,pos+0,zero,pos+3+1, dates[0], left_border)
    worksheet.merge_range(zero,pos+4+1,zero,pos+7+2, dates[1], left_border)
    worksheet.merge_range(zero,pos+4+4+2,zero,pos+7+4+3, dates[2], left_border)
    worksheet.merge_range(zero,pos+4+4+4+3,zero,pos+7+4+4+4, dates[3], left_border)
    worksheet.merge_range(zero,pos+4+4+4+4+4,zero,pos+7+4+4+4+5, dates[4], left_border)
    worksheet.merge_range(zero,pos+4+4+4+4+4+5,zero,pos+7+4+4+4+4+6, dates[5], left_border)

    par=('1-2','3-4','5-6','7-8','9-10')
    k=0
    for i in range(1,31):
        if k==0:
            worksheet.write(zero+1,i,par[k],left_border)
        else:
            worksheet.write(zero+1,i,par[k])
        if k==4:
            k=0
        else:
            k+=1

    for i in range(names*4):
        worksheet.write(zero+i+2,0,'',right_border)
        worksheet.write(zero+i+2,1,'',left_border)
        worksheet.write(zero+i+2,5,'',right_border)
        worksheet.write(zero+i+2,6,'',left_border)
        worksheet.write(zero+i+2,10,'',right_border)
        worksheet.write(zero+i+2,11,'',left_border)
        worksheet.write(zero+i+2,15,'',right_border)
        worksheet.write(zero+i+2,16,'',left_border)
        worksheet.write(zero+i+2,20,'',right_border)
        worksheet.write(zero+i+2,21,'',left_border)
        worksheet.write(zero+i+2,25,'',right_border)
        worksheet.write(zero+i+2,26,'',left_border)

    for i in range(names):
        for j in range(31):
            if j==5 or j==10 or j==15 or j==20 or j==25:
                worksheet.write((zero+i)*4+2,j,'',top_right_border)
            elif j==6 or j==11 or j==16 or j==21 or j==26:
                worksheet.write((zero+i)*4+2,j,'',top_left_border)
            else:
                worksheet.write((zero+i)*4+2,j,'',top_border)

    return workbook


def write_name(workbook,name,man):
    border=workbook.add_format({'right':5,'top':1})

    worksheet=workbook.worksheets()[0]
    pos_row = man*4+2
    worksheet.merge_range(pos_row,0,pos_row+3,0, name, border)
    return workbook

def write_data(workbook,paras,day=0,man=0,color_set=0):
    worksheet=workbook.worksheets()[0]

    pos_st = day*5+1
    pos_row = man*4+2

    #Установка цвета для дня
    colors=['#0000FF','#00FFFF','#800000','#008000','#00FF00','#FF00FF','#000080','#FFFF00','#800080','#00FF00','#FF6600','#FF0000']
    color = workbook.add_format()

    colored_top_border=workbook.add_format({'top':1})
    color.set_pattern(1)
    if color_set==0 or color_set+1>len(colors):
        colored_top_border.set_bg_color('#FFFFFF')
        color.set_bg_color('#FFFFFF')
    else:
        color.set_bg_color(colors[color_set-1])
        colored_top_border.set_bg_color(colors[color_set-1])

    for i in range(5):
        for j in range(len(paras)):
            if paras[j][0]==i+1:
                i=pos_st+i
                worksheet.write(pos_row,i,paras[j][1],colored_top_border)
                worksheet.write(pos_row+2,i,paras[j][2],color)
                worksheet.write(pos_row+3,i,paras[j][4].replace('.0',''),color)
                worksheet.write(pos_row+1,i,paras[j][5].replace('.0',''),color)

    return workbook

def save(workbook):
    workbook.close()