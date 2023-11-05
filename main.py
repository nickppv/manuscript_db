import openpyxl
import sqlite3
import csv


class DataBase:

    def __init__(self, name):
        self.name = name

    def create_db(self):
        '''создаем БД и таблицы'''

        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS country (
                id_country INTEGER PRIMARY KEY,
                name_country VARCHAR(40) UNIQUE)'''
                )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS isg (
                id_isg INTEGER PRIMARY KEY UNIQUE,
                name_isg TEXT)'''
                )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS goods (
                id_tovar INTEGER PRIMARY KEY NOT NULL UNIQUE,
                name_tovar TEXT,
                barcod TEXT,
                id_country INTEGER,
                id_isg INTEGER,
                FOREIGN KEY (id_country) REFERENCES country (id_country),
                FOREIGN KEY (id_isg) REFERENCES isg (id_isg))'''
                )
        cursor.close()
        conn.close()

    def fill_db(self, exl_file):
        '''заполняем БД данными'''

        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        file = openpyxl.open(exl_file)
        sheet = file.active
        for i in range(2, sheet.max_row+1):
            dump = tuple()
            for j in range(0, 6):
                dump += (sheet[i][j].value, )
            cursor.execute(
                '''INSERT OR IGNORE INTO country (name_country)
                VALUES (?)''', (dump[4], )
                )
            cursor.execute(
                '''INSERT OR IGNORE INTO isg (id_isg, name_isg)
                VALUES (?, ?)''', (dump[2], dump[3])
                )
            cursor.execute(
                '''INSERT OR IGNORE INTO goods (
                    id_tovar, name_tovar, barcod, id_country, id_isg) VALUES (?, ?, ?, (
                    SELECT id_country FROM country WHERE name_country = ?), (
                    SELECT id_isg FROM isg WHERE name_isg = ?)
                )''', (str(dump[0]).lstrip('--'), str(dump[1]), dump[-1], dump[4], dump[3])
                )
            conn.commit()
        file.close()
        cursor.close()
        conn.close()

    def get_goods(self):
        '''получаем количество товаров по странам'''

        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        cursor.execute('''SELECT country.name_country AS name_country,
                       COUNT(goods.id_country) AS cnt FROM country
                       INNER JOIN goods ON (country.id_country =
                       goods.id_country) GROUP BY name_country''')
        result = cursor.fetchall()
        cursor.close()
        conn.close()
        return result


class TSV_File:

    def __init__(self, inner_data):
        self.inner_data = inner_data

    def get_result(self):
        '''вывод табличного файла'''

        with open('data.tsv', 'w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file, delimiter='\t')
            for row in self.inner_data:
                writer.writerow(row)


database = DataBase('base.sqlite')
database.create_db()
database.fill_db('data.xlsx')
out_data = database.get_goods()
data_tsv = TSV_File(out_data)
data_tsv.get_result()
