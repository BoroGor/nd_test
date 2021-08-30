from openpyxl import load_workbook
import PySimpleGUI as sg
from random import choice


# загружаю табличку с вопросами и ответами
wb = load_workbook('try.xlsx')

# создаю переменную активного листа (по умолчанию Лист1)
ws1 = wb.active

# цикл подсчёта непустых строк
def count_of_row(worksh):
    row_m = 1
    while True:
        if worksh[f'A{row_m}'].value != None:
            row_m = row_m + 1
        else:
            break
    # сохраняем номер последней непустой строки
    return(row_m - 1)

# записываем номер последней непустой строки
num_of_rows = count_of_row(ws1)

# список с номерами строк
ln = [x for x in range(1, num_of_rows + 1)]
# функция обновления текстовой части вопроса
def update_q(window, text_key, row_max, worksh, list_num):
    """window - активное окно, text_key - идентификатор текстового поля,
    row_max - число вопросов, worksh - лист книги Excell с вопросами,
    list_num - список, состоящий из номеров вопросов.
    """
    # если список номеров пустой
    if len(list_num) == 0:
        # заполняем лист заново
        list_num = [x for x in range(1, row_max + 1)]
    # выбираем случайный номер вопроса
    r = choice(list_num)
    # убираем из списка выбранный элемент (номер вопроса)
    list_num.remove(r)
    # получаем значение ячейки - вопрос (номер вопроса = номеру строки)
    val = worksh[f'A{r}'].value
    # обновление значения текстового поля, вывод нового вопроса
    window[text_key].update(val)
    # возвращает [номер вопроса (строки), уменьшенный список номеров вопросов]
    return r, list_num

# функция обновления ответа (изображения)
def update_a(window, img_key, worksh, row):
    """window - активное окно, img_key - ключ поля с изображением,
    worksh - лист книги Excell с вопросами, row - номер строки (вопроса,ответа)
    """
    # запись названия файла из соответствующей ячейки таблицы
    filename = worksh[f'B{row}'].value
    # обновление изображения
    window[img_key].update(filename=filename)

# функция обновления текстового поля с числом оставшихся вопросов
def update_len(window, text_key, l):
    """window - активное окно, text_key - ключ текстового поля, l - список
    номеров вопросов
    """
    # обновление текстового поля
    window[text_key].update(len(l))

#----------Опишем внешний вид окна----------#
layout = [
    # 1 строка - текст вопроса
    [sg.Text(120*'#', key="-qtext-")],
    # 2 строка - кнопка обновления вопроса
    [sg.Button('Вопрос', enable_events=True, key="-qbut-")],
    # 3 строка - кнопка обновления ответа
    [sg.Button('Ответ', enable_events=True, key="-abut-")],
    # 4 строка - изображение
    [sg.Image(r'C:\Users\Egor\projects\nd_test\фон.png', key="-img-")],
    # 5 строка - число оставшихся вопросов
    [sg.Text(str(len(ln)), key="-len-", font="Arial, 18")]
]

# создание объекта окна с названием и расположением элементов
window = sg.Window('Тест по Надёжности', layout, size=(1050,800))

#-----Описание событий-----#
while True:
    # получение данных о событиях в окне
    event, values = window.read()
    # если событие совпадает с одним из кортежа (закрытие окна)
    if event in (sg.WIN_CLOSED, 'Exit'):
        # прервать работу
        break
    # если событие соответствует нажатию кнопки обновления вопроса
    elif event == "-qbut-":
        # обновление переменной, отвечающей за номер вопроса (строки и ответа),
        # обновление списка номеров вопросов
        rn, ln = update_q(window, "-qtext-", num_of_rows, ws1, ln)
        # обновление поля с числом оставшихся вопросов
        update_len(window, "-len-", ln)
    # если событие соответствует нажатию кнопки обновления ответа
    elif event == "-abut-":
        # обновление изображения
        update_a(window, "-img-", ws1, rn)

# Закрываю окно
window.close()
