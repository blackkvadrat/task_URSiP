# импортируемый необходимые библиотеки
import sqlite3
import openpyxl

# создаем базу данных
conn = sqlite3.connect('URSiP.db')
# создаем объект курсора для работы с базой данных
cursor = conn.cursor()
# создаем таблицу
cursor.execute('''CREATE TABLE IF NOT EXISTS mytable  (
                  id INTEGER PRIMARY KEY AUTOINCREMENT,
                  company TEXT,
                  fact_qliq_data1 INTEGER,
                  fact_qliq_data2 INTEGER,
                  fact_qoil_data1 INTEGER,
                  fact_qoil_data2 INTEGER,
                  forecast_qliq_data1 INTEGER,
                  forecast_qliq_data2 INTEGER,
                  forecast_qoil_data1 INTEGER,
                  forecast_qoil_data2 INTEGER)''')
# открываем xlsx файл
book = openpyxl.open("attachment.xlsx", read_only=True)
sheet = book.active
# читаем данные из таблицы
for row in range(4, sheet.max_row + 1):
    company = sheet[row][1].value
    fact_qliq_data1 = int(sheet[row][2].value)
    fact_qliq_data2 = int(sheet[row][3].value)
    fact_qoil_data1 = int(sheet[row][4].value)
    fact_qoil_data2 = int(sheet[row][5].value)
    forecast_qliq_data1 = int(sheet[row][6].value)
    forecast_qliq_data2 = int(sheet[row][7].value)
    forecast_qoil_data1 = int(sheet[row][8].value)
    forecast_qoil_data2 = int(sheet[row][9].value)
# записываем данные из таблицы в базу данных
    cursor.execute("INSERT INTO mytable (company, fact_qliq_data1, fact_qliq_data2, fact_qoil_data1, \
    fact_qoil_data2, forecast_qliq_data1,forecast_qliq_data2, forecast_qoil_data1, forecast_qoil_data2) \
    VALUES (?,?,?,?,?,?,?,?,?)", (company, fact_qliq_data1, fact_qliq_data2, fact_qoil_data1,
                                    fact_qoil_data2, forecast_qliq_data1, forecast_qliq_data2, forecast_qoil_data1,
                                    forecast_qoil_data2))
# считаем total по каждой колонке и сохраняем в список totals
totals = cursor.execute("SELECT SUM(fact_qliq_data1), SUM(fact_qliq_data2), SUM(fact_qoil_data1), \
                         SUM(fact_qoil_data2), SUM(forecast_qliq_data1), SUM(forecast_qliq_data2), \
                         SUM(forecast_qoil_data1), SUM(forecast_qoil_data2) FROM mytable").fetchone()

# добавляем новую строку с полученными суммами в таблицу
cursor.execute("INSERT INTO mytable (company, fact_qliq_data1, fact_qliq_data2, fact_qoil_data1, \
              fact_qoil_data2, forecast_qliq_data1, forecast_qliq_data2, forecast_qoil_data1, \
              forecast_qoil_data2) VALUES ('total', ?, ?, ?, ?, ?, ?, ?, ?)",
               totals)

# сохраняем изменения и закрываем подключение
conn.commit()
conn.close()
