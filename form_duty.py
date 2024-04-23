import openpyxl
import datetime

"""
Скрипт для формирования раскладок в соответствии с графиком дежурств. Автор Панин С.А.
Как пользоваться
!!! Внимание !!! Не допускаются изменения структуры шаблона, без согласования его с данным скриптом. Так нельзя менять
 формулы и добавлять, удалять строки или столбцы, так как это приведете к изменению координат, заданных в скрипте
1. Заполнить шаблон в файле "Исходный_Шаблон.xlsx"
    - Заполнить количество участников на каждый день
    - Заполнить кто, когда дежурит
    - Заполнить базовые блюда в колонке блюд
    - Провести пересчет ячейки "Общее кол-во приемов пищи"
    - Если требуется добавить по аналогии новые блюда
2. Согласовать скрипт с шаблоном
    - Заполнить словарь MEAL_PLACE в соответствии с координатами количества блюд
    - Заполнить переменные start_x, end_x, start_y, end_y в соответствии размерами таблицы дежурств
    - Заполнить переменную out_file_name в соответствии с именем итогового файла дежурств
3. Выполнить скрипт python form_duty.py. Естественно для этого потребуется установленный python и пакет openpyxl
4a. Если скрипт выполнился без ошибок вручную проверить итоговую раскладку. Если обнаружены ошибки обратиться к 
    разработчику за разъяснениями
4b. Если скрипт выполнился с ошибками обратиться к разработчику за разъяснениями
"""


def get_cell_x_num_by_letter(letter):
    if ord(letter) >= ord('A') and ord(letter) <= ord('Z'):
        return ord(letter) - ord('A') + 1
    elif ord(letter) >= ord('a') and ord(letter) <= ord('z'):
        return ord(letter) - ord('a') + 1
    else:
        raise Exception("<Error> Wrong letter {}".format(letter))


def get_time_of_day_str(tod):
    if tod == 0:
        return "Завтрак"
    elif tod == 1:
        return "Обед"
    elif tod == 2:
        return "Ужин"
    else:
        raise Exception("<Error> wrong time of date {}".format(tod))


class FoodTime(object):
    def __init__(self, person, date, time_of_day, general_food_type, meal_portions):
        self.person = person
        self.date = date
        self.time_of_day = time_of_day
        self.general_food_type = general_food_type
        self.meal_portions = meal_portions

    def __repr__(self):
        return "{} {} {} {} {} порций".format(
            self.date.strftime('%Y-%m-%d'),
            self.person,
            self.time_of_day,
            self.general_food_type,
            self.meal_portions
        )


# Словарь с расположением типов еды на листе
MEAL_PLACE = {
    "Геркулес": "G8", "Пшенка": "I8",
    "Харчо": "G14", "Борщ": "I14", "Гороховый": "K14",
    "Гречка": "G20", "Рис": "I20", "Макароны": "K20"
}


def form_duty(patten_file, out_file_name):
    wb = openpyxl.load_workbook(patten_file, read_only=False)  # Открыть таблицу с данными о НС
    sheet_names = wb.sheetnames
    print("Найденные листы в книге: {}".format(sheet_names))
    sheet_name = "График"
    ws = wb[sheet_name]

    #some = ws.cell(row=10, column=6).value
    #print(some)
    #quit()

    start_x = get_cell_x_num_by_letter('F')
    end_x = get_cell_x_num_by_letter('N')
    print("start_x = {}, end_x = {}".format(start_x, end_x))
    meal_x = get_cell_x_num_by_letter('E')      # Колонка с блюдом
    meal_count_x = get_cell_x_num_by_letter('D')  # Колонка с количеством персон

    start_y = 8
    end_y = 34

    all_count = 0

    data_base = []
    start_date = datetime.datetime(year=2024, month=5, day=4)
    for x in range(start_x, end_x + 1):
        date = start_date + datetime.timedelta(days=x - 1)
        for y in range(start_y, end_y + 1):
            person = ws.cell(row=y, column=x).value
            if person is None:
                continue
            print("row = {} col = {} val = {}".format(y, x, person))
            time_of_day = (y - start_y) % 3
            meal = ws.cell(row=y, column=meal_x).value
            meal_portions = ws.cell(row=y, column=meal_count_x).value
            data_base.append(FoodTime(
                person=person,
                date=date,
                time_of_day=get_time_of_day_str(time_of_day),
                general_food_type=meal,
                meal_portions=meal_portions
            ))
            all_count += 1

    # FixMe добавить вычисление формулы подсчета количества дежурств
    test_count = int(ws['P6'].value)  # Проверяем данные

    print("all_count = {}, test_count = {}".format(all_count, test_count))
    if all_count != test_count:
        raise Exception("all_count != test_count {} != {}".format(all_count, test_count))

    i = 0
    for d in data_base:
        print("{}. {}".format(i, d))
        i += 1

    persons = set([ft.person for ft in data_base])   # Получаем имена всех участников
    persons = sorted(persons)
    persons_meals = {}
    for p in persons:
        target = wb.copy_worksheet(wb["Шаблон"])
        target.title = p
        persons_meals[p] = {}
        for ft in data_base:
            if ft.person == p:
                try:
                    persons_meals[p][ft.general_food_type] += ft.meal_portions
                except KeyError:
                    persons_meals[p][ft.general_food_type] = ft.meal_portions

        for meal in persons_meals[p]:
            target[MEAL_PLACE[meal]] = persons_meals[p][meal]


    print(persons_meals)

    wb.save(out_file_name)


form_duty(patten_file="Исходный_Шаблон.xlsx", out_file_name="Раскладка_Питание_24.05.04.xlsx")


