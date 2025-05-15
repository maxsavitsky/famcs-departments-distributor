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

    def __lt__(self, other):
        if self.score == other.score:
            return self.full_name < other.full_name
        return self.score > other.score

def column_letter_to_index(letter: str) -> int:
    return ord(letter.upper()[0]) - ord('A') + 1

class DepartmentInfo:
    def __init__(self, students_limit: int, results_column: str):
        self.students_limit = students_limit
        self.results_column = results_column

    def get_results_column_index(self):
        return column_letter_to_index(self.results_column)

spreadsheet = gspread.service_account(filename='credentials.json').open_by_url("https://docs.google.com/spreadsheets/d/1_Qq95tMhy7mlyoYDWhpIQ7HgNZo0QeIPELi9ISYQecI/")
main_ws = spreadsheet.get_worksheet(0)
results_ws = spreadsheet.get_worksheet(1)

column_for_not_distributed = column_letter_to_index('Y')

departments: dict[str, DepartmentInfo] = {
    "МСС":   DepartmentInfo(12, 'A'),
    "ТП":    DepartmentInfo(13, 'D'),
    "ИСУ":   DepartmentInfo(13, 'G'),
    "КТС-2": DepartmentInfo(12, 'J'),
    "ДМА":   DepartmentInfo(13, 'M'),
    "БМИ":   DepartmentInfo(12, 'P'),
    "КТС-4": DepartmentInfo(12, 'S'),
    "ФМиИС": DepartmentInfo(12, 'V'),
}
group_boundaries = [
    (1, 26),
    (29, 53),
    (56, 80),
    (83, 105)
]

def to_a1(row1, col1, row2, col2):
    return gspread.utils.rowcol_to_a1(row1, col1) + ":" + gspread.utils.rowcol_to_a1(row2, col2)

def parse_users() -> list[Student]:
    all_values = main_ws.get_all_values()

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
                print(f"Could not convert score ({all_values[i][3]}) to float. {i}")
            students += [Student(full_name, score, priors, group_num)]

    return students


def distribute(students: list[Student]):
    deps = {}
    for dep in departments.keys():
        deps[dep] = list()

    students.sort()
    for s in students:
        print(s.score, s.full_name)

    for student in students:
        for priority in range(8):
            prior = student.priorities[priority]
            if prior not in deps:
                continue
            if len(deps[prior]) < departments[prior].students_limit:
                deps[prior].append(student)
                student.assigned_department = prior
                student.assigned_by_which_priority = priority + 1
                break

    for dep_name, st in deps.items(): # type: (str, list[Student])
        dep_col = departments[dep_name].get_results_column_index()
        results_ws.batch_clear([to_a1(3, dep_col, 50, dep_col + 1)])
        if len(st) == 0:
            continue
        data = [[f"{s.full_name} ({s.assigned_by_which_priority})", str(s.score)] for s in st]
        print(dep_name)
        print(*data)
        results_ws.update(data, to_a1(3, dep_col, 2 + len(data), dep_col + 1))

    not_distributed = [s for s in students if s.assigned_department is None]
    not_distributed.sort()
    results_ws.batch_clear([to_a1(3, column_for_not_distributed, 100, column_for_not_distributed + 1)])
    if len(not_distributed) > 0:
        data = [[s.full_name, str(s.score)] for s in not_distributed]
        results_ws.update(data, to_a1(3, column_for_not_distributed, 2 + len(data), column_for_not_distributed + 1))
    results_ws.update_cell(1, 1, f"Время обновления: {str(datetime.datetime.today())}")


if __name__ == '__main__':
    students = parse_users()
    print(len(students))
    distribute(students)

