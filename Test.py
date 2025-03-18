

default_name = {1: "sorted_by_number_task.xlsx",
                2: "sorted_by_diag_type.xlsx"}


def test_files(xlsx_file):
    if not xlsx_file.lower().endswith(".xlsx"):
       xlsx_file += ".xlsx"
    if not xlsx_file:
       print("Файл не найден! Повторите попытку")
    return xlsx_file