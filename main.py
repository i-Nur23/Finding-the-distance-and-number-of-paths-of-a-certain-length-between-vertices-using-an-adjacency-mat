from openpyxl import load_workbook
import numpy as np

matrix = []

def FindWay(current_matrix,v1,v2,count):
    way = 1
    while (current_matrix[v1-1][v2-1] == 0):
        current_matrix = np.dot(current_matrix,matrix)
        way += 1
        if (way == count):
            return -1
    return way




def main():
    current_matrix = matrix
    count = len(current_matrix)
    while (True):
        choice = int(input("Выберите действие:\n"
                           "1 - найти расстояние между вершинами\n"
                           "2 - найти количество маршрутов определённой длины между вершинами\n"
                           "Выбор: "))
        if (choice != 1) and (choice != 2):
            print("Введите 1 или 2 !!!")
            continue
        print("")
        if (choice == 1):
            v1 = int(input("Первая вершина: "))
            v2 = int (input("Вторая вершина: "))
            if (v1 > count) or (v2 > count) or (v1 < 1) or (v2 < 1):
                print("Такой вершины не существует!!!")
                continue
            way = FindWay(current_matrix,v1,v2,count)
            if (way == -1):
                print("Вершины v"+str(v1)+" и v"+str(v2)+" не связаны")
            else:
                print("Расстояние: "+str(way))
            print("")
        while (True):
            close = int(input("Выберите действие:\n"
                               "1 - ещё раз\n"
                               "2 - закрыть\n"
                               "Выбор: "))
            if (close != 1) and (close != 2):
                print("Введите 1 или 2 !!!")
                continue
            break
        if (close == 2):
            break



def ReadFile(count):
    global matrix
    row = []
    wb = load_workbook("Matrix.xlsx")
    sheet = wb['Matrix']
    if (sheet.max_row != sheet.max_column):
        print("Матрица должна быть квадратной!!!")
        return False

    if (sheet.max_row != count+1):
        print("Кол-во вершин в матрице не совпадает с указанным!!!\n"
              "Первые столбец и строка зарезервированы для вершин!!!")
        return False

    for i in range (count):
        for j in range (count):
            value = sheet.cell(row=2 + i, column=2 + j).value
            if (value != None) and (type(value) == int ) and (value > -1):
                row.append(sheet.cell(row = 2 + i, column = 2 + j).value)
            else:
                print("В таблице имеется пустая/не целочисленная/отрицательная ячейка!!!")
                return False
        matrix.append(row)
        row = []
    return True

while (True):
    vertex_count = int(input("Введите количестов вершин графа: "))
    if (vertex_count < 1):
        print("Количестов строк не может быть отрицательным числом!!!")
        continue
    break

if (ReadFile(vertex_count)):
    main()



