from openpyxl import load_workbook

matrix = []

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
    print("Файл прочитан успешно")



