import sqlite3
import openpyxl


# создаем базу данных, таблицу
class Table:
    def __init__(self, db_name: str, table_name: str):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.table_name = table_name

    def create_table(self):
        self.cursor.execute(f'''CREATE TABLE IF NOT EXISTS {self.table_name} (
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
        self.conn.commit()

    # метод для вставки данных в таблицу
    def insert_data(self, company, fact_qliq_data1, fact_qliq_data2, fact_qoil_data1,
                    fact_qoil_data2, forecast_qliq_data1, forecast_qliq_data2,
                    forecast_qoil_data1, forecast_qoil_data2):
        self.cursor.execute(f"INSERT INTO {self.table_name} (company, fact_qliq_data1, fact_qliq_data2,\
                     fact_qoil_data1, fact_qoil_data2, forecast_qliq_data1,forecast_qliq_data2, forecast_qoil_data1,\
                    forecast_qoil_data2) VALUES (?,?,?,?,?,?,?,?,?)", (company, fact_qliq_data1, \
                                                                       fact_qliq_data2, fact_qoil_data1, \
                                                                       fact_qoil_data2, forecast_qliq_data1, \
                                                                       forecast_qliq_data2, forecast_qoil_data1,
                                                                       forecast_qoil_data2))
        self.conn.commit()  # после каждой манипуляции с бд по вставке или удалению данных сохраняем их

    # суммируем значения
    def select_summ(self):
        totals = self.cursor.execute(f"SELECT SUM(fact_qliq_data1), SUM(fact_qliq_data2), SUM(fact_qoil_data1), \
                         SUM(fact_qoil_data2), SUM(forecast_qliq_data1), SUM(forecast_qliq_data2), \
                         SUM(forecast_qoil_data1), SUM(forecast_qoil_data2) FROM {self.table_name}").fetchone()
        return totals

    # вставляем значения в таблицу
    def insert_sum(self):
        totals = self.select_summ()
        self.cursor.execute(f"INSERT INTO {self.table_name} (company, fact_qliq_data1, fact_qliq_data2, fact_qoil_data1, \
              fact_qoil_data2, forecast_qliq_data1, forecast_qliq_data2, forecast_qoil_data1, \
              forecast_qoil_data2) VALUES ('total', ?, ?, ?, ?, ?, ?, ?, ?)",
                            totals)

        self.conn.commit()  # сохраняем изменения
        self.conn.close()  # закрываем соединения


if __name__ == "__main__":
    table = Table('URSiP2.db', 'mytable')
    table.create_table()

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
        # записываем значения из файда в базу данных
        table.insert_data(company, fact_qliq_data1, fact_qliq_data2, fact_qoil_data1,
                          fact_qoil_data2, forecast_qliq_data1, forecast_qliq_data2,
                          forecast_qoil_data1, forecast_qoil_data2)
    # вычисляем тотал
    table.insert_sum()
