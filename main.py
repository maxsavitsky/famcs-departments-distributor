import gspread
import datetime

class Student:
    def __init__(self, full_name: str, score: float, priorities: list[str], group_number: int):
        self.full_name = full_name
        self.score = score
        self.priorities = priorities
        self.group_number = group_number
        self.assigned_department = None
        self.assigned_by_which_priority = -1

spreadsheet = gspread.service_account(filename='credentials.json').open_by_url("https://docs.google.com/spreadsheets/d/1_Qq95tMhy7mlyoYDWhpIQ7HgNZo0QeIPELi9ISYQecI/")
main_ws = spreadsheet.get_worksheet(0)
results_ws = spreadsheet.get_worksheet(1)

departments = {
    "ДМА": (13, 12), # (limit, column)
    "МСС": (12, 0),
    "ТП": (13, 3),
    "ИСУ": (13, 6),
    "КТС-2": (12, 9),
    "КТС-4": (12, 18),
    "БМИ": (12, 15),
    "ФМиИС": (12, 21),
}

def to_a1(row1, col1, row2, col2):
    return gspread.utils.rowcol_to_a1(row1, col1) + ":" + gspread.utils.rowcol_to_a1(row2, col2)

def parse_users() -> list[Student]:
    all_values = main_ws.get_all_values()
    group_boundaries = [
        (1, 26),
        (29, 53),
        (56, 80),
        (83, 105)
    ]

    students = []

    for group_num in range(len(group_boundaries)):
        start = group_boundaries[group_num][0]
        end = group_boundaries[group_num][1]

        for i in range(start, end + 1):
            priors = all_values[i][7:]
            full_name = all_values[i][1]
            score = 0.0
            try:
                score = float(all_values[i][3].replace(',', '.'))
            except ValueError:
                print(f"Could not convert score ({all_values[i][3]}) to float. Skipping {i}")
            students += [Student(full_name, score, priors, group_num)]

    return students


def distribute(students: list[Student]):
    deps = {}
    for dep in departments.keys():
        deps[dep] = list()

    students.sort(key=lambda s: s.score, reverse=True)

    for student in students:
        for priority in range(8):
            prior = student.priorities[priority]
            if prior not in deps:
                continue
            if len(deps[prior]) < departments[prior][0]:
                deps[prior].append(student)
                student.assigned_department = prior
                student.assigned_by_which_priority = priority + 1
                break

    for dep_name, st in deps.items(): # type: (str, list[Student])
        dep_col = departments[dep_name][1]
        results_ws.batch_clear([to_a1(3, dep_col + 1, departments[dep_name][0] + 2, dep_col + 2)])
        if len(st) == 0:
            continue
        data = [[f"{s.full_name} ({s.assigned_by_which_priority})", str(s.score)] for s in st]
        print(dep_name)
        print(*data)
        results_ws.update(data, to_a1(3, dep_col + 1, 2 + len(data), dep_col + 2))

    not_distributed = [s for s in students if s.assigned_department is None]
    not_distributed.sort(key=lambda s: s.score, reverse=True)
    col_not_distributed = 24
    results_ws.batch_clear([to_a1(3, col_not_distributed, 100, col_not_distributed + 1)])
    if len(not_distributed) > 0:
        data = [[s.full_name, str(s.score)] for s in not_distributed]
        results_ws.update(data, to_a1(3, col_not_distributed + 1, 2 + len(data), col_not_distributed + 2))
    results_ws.update_cell(1, 1, f"Время обновления: {str(datetime.datetime.today())}")


if __name__ == '__main__':
    students = parse_users()
    print(len(students))
    distribute(students)

