import struct
import openpyxl
from openpyxl.styles import Alignment
from Global import manual_widths

def parse_bin_file(bin_file):
    headers = ["Порядковый номер", "Время", "Номер задачи", "Тип диагностического сообщения",
               "Длина бинарных данных", "Бинарные данные", "Текстовое сообщение разработчику"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Данные 1"
    ws.append(headers)

    # Устанавливаем ширину столбцов, если manual_widths существует
    for i, width in manual_widths.items():
        col_letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[col_letter].width = width

    for col in range(4, 6):
        ws.cell(row=1, column=col).alignment = Alignment(wrap_text=True)

    sheet_number = 1
    row_count = 1

    try:
        with open(bin_file, "rb") as f:
            while True:
                header = f.read(12)
                if not header or len(header) < 12:
                    break

                sequence_number, time_val, task_number, diag_type, bin_data_length = struct.unpack("<IIBBH", header)
                binary_data = f.read(bin_data_length) if bin_data_length > 0 else b""
                bin_data_hex = binary_data.hex().upper() if binary_data else "Отсутствует"
                
                # Чтение текстового сообщения разработчику
                dev_mess = bytearray()
                while True:
                    byte = f.read(1)
                    if not byte or byte == b'\0':
                        break
                    dev_mess.extend(byte)

                try:
                    dev_message = dev_mess.decode("koi8-r")
                except UnicodeDecodeError:
                    dev_message = dev_mess.decode("koi8-r", errors="replace")

                # Добавляем строку в текущий лист
                ws.append([sequence_number, time_val, task_number, diag_type, bin_data_length, bin_data_hex, dev_message])
                row_count += 1

                # Если строк больше 10000, создаем новый лист
                if row_count > 10000:
                    sheet_number += 1
                    ws = wb.create_sheet(title=f"Данные {sheet_number}")
                    ws.append(headers)
                    # Устанавливаем ширину столбцов для нового листа
                    for i, width in manual_widths.items():
                        col_letter = openpyxl.utils.get_column_letter(i)
                        ws.column_dimensions[col_letter].width = width
                    for col in range(4, 6):
                        ws.cell(row=1, column=col).alignment = Alignment(wrap_text=True)
                    row_count = 1  # Сброс счетчика строк для нового листа

    except Exception as e:
        print(f"Ошибка при чтении BIN файла: {e}")
        return None

    # Устанавливаем выравнивание для всех ячеек
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row=2, min_col=6, max_col=7):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)

    return wb
