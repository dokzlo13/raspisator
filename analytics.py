__author__ = 'zloy'

def find_prep(massive, name, day):
    """
    Функция поиска предметов, которые есть у преподавателя в указанный день
    :param massive: принимает расписание по курсам
    :param name: принимает имя преподавателя в виде строки или кортежа из возможных имен, встречающихся в расписании
    :param day: день, в котором нужно найти пары указанного преподавателя
    :return: возращает массив с парами
    """
    mass = []
    for i in range(len(massive)):
        for j in range(len(massive[i][day])):
            try:
                if type(name) == tuple:
                    for k in name:
                        if massive[i][day][j][3] == k:
                            mass.append(massive[i][day][j])
                else:
                    if massive[i][day][j][3] == name:
                        mass.append(massive[i][day][j])
            except IndexError:
                pass
    return mass

def prep_conn_groups(massive):
    for i in range(len(massive)):
        for j in range(len(massive[0])):
            for k in range(20):
                try:
                    print(massive[i][j][k])
                except IndexError:
                    pass


def find_class(massive, group, day):
    mass = []
    for i in range(len(massive)):
        for j in range(len(massive[i][day])):
            try:
                if massive[i][day][j][5] == group:
                    mass.append(massive[i][day][j])
            except IndexError:
                pass
    return mass

def find_auditor(massive, auditor, day):
    mass = []
    for i in range(len(massive)):
        for j in range(len(massive[i][day])):
            try:
                if massive[i][day][j][4] == auditor:
                    mass.append(massive[i][day][j])
            except IndexError:
                pass
    return mass
