import os.path

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import pyodbc

import pandas as pd
import configparser

import logging
import logging.config
import time
import io
from PIL import Image
import ipaddress


class Winkel:

    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read('/home/exact/PycharmProjects/Winkel/Winkel.ini')
        logging.config.fileConfig('/home/exact/PycharmProjects/Winkel/Winkel.conf')
        self.logger = logging.getLogger('WinkelLogger')
        self.exact_engine = None
        self.exact_connection = None
        self.exact_Session = None
        self.exact_session = None

        self.ds21_engine, self.ds21_connection = get_engine(self.logger, self.config, 'ds21')
        self.ds21_Session = sessionmaker(bind=self.ds21_engine, autocommit=True, autoflush=True)
        self.ds21_session = self.ds21_Session()

    def ms_to_my(self):
        self.exact_engine, self.exact_connection = get_engine(self.logger, self.config, 'exact')
        self.exact_Session = sessionmaker(bind=self.exact_engine, autocommit=True, autoflush=True)
        self.exact_session = self.exact_Session()
        tables = self.config['TABLES']
        for table in tables:
            self.logger.info('Table: {0}'.format(table))
            query = tables[table]
            self.logger.info(query)
            df = pd.read_sql_query(query, self.exact_connection)
            df.to_sql(name=table, con=self.ds21_engine, if_exists='replace')
            self.logger.info('Resultaat query overgezet naar lokale mysql db')

        self.exact_connection.close()
        self.exact_engine.dispose()

        self.logger.info('ms_to_my finished.')

    def products(self):
        header = 1
        directory = self.config['FILES']['directory']
        self.logger.info(directory)
        sql = self.config['FILES']['products']
        self.logger.info(sql)
        # Connect to the database
        cursor = self.ds21_connection.cursor()
        cursor.execute(sql)
        cols = fields(cursor)
        with open(os.path.join(directory, 'products_data.csv'), 'w') as prod:
            with open(os.path.join(directory, 'picture_data.csv'), 'w') as pict:
                rows = cursor.fetchall()
                for row in rows:
                    self.logger.info("row : %s", row[0])
                    if header == 1:
                        header = 0
                        self.logger.info("pict.write(row[0]) %s", row[0])
                        self.logger.info(row[cols['header_prod1']])
                        self.logger.info(row[cols['header_pict1']])
                        pict.write(row[cols['header_pict1']] + "\n")
                        pict.write(row[cols['header_pict2']] + "\n")

#                        self.logger.info(pict.write(cols["header_prod1"]))
                        prod.write(row[cols['header_prod1']] + "\n")
                        prod.write(row[cols['header_prod2']] + "\n")

                    if row[cols['picture']] == 'Y':
                        pict.write(row[cols["SearchCode"]] + row[cols["detailed_remote"]] + "\n")
                        prod.write(row[cols['fields']] + row[cols['thumb_remote']] + row[cols['detailed_remote']] + "\n")
                    else:
                        prod.write(row[cols['fields']] + "\n")

    def images(self):
        # Opens a image in RGB mode
        directory = self.config['PICTURES']['directory']
        self.logger.info(directory)
        sql = self.config['PICTURES']['sql']
        self.logger.info(sql)
        self.exact_engine, self.exact_connection = get_engine(self.logger, self.config, 'exact')
        self.exact_Session = sessionmaker(bind=self.exact_engine, autocommit=True, autoflush=True)
        cursor = self.exact_connection.cursor()
        cursor.execute(sql)
        cols = fields(cursor)
        new_size = (int(self.config['PICTURES']['thumb_width']), int(self.config['PICTURES']['thumb_height']))
        thumb_directory = self.config['PICTURES']['thumb_directory']
        for row in cursor.fetchall():
            image = Image.open(io.BytesIO(row[1]))
            # Naam moet overeenkomen met naam in url in views!
            filename = f"pict_{row[cols['id']]}.jpg"
            image.save(os.path.join(directory, filename))
            im = Image.open(os.path.join(directory, filename))
            im1 = im.resize(new_size)
            # Naam moet overeenkomen met naam in url in views!
            filename = f"thumb_{row[cols['id']]}.jpg"
            im1.save(os.path.join(thumb_directory, filename))

        self.exact_connection.close()
        self.exact_engine.dispose()

    def xcart_files(self):
        # Connect to the database
        cursor = self.ds21_connection.cursor()
        for step in self.config['FILES']:
            if step == 'directory':
                directory = self.config['FILES']['directory']
            elif step == 'products':
                self.products()
            else:
                sql = self.config['FILES'][step]
                file = os.path.join(directory, f"{step}.csv")
                self.logger.info(sql)
                cursor.execute(sql)
                cols = fields(cursor)
                header = 1
                with open(file, 'w') as ws:
                    rows = cursor.fetchall()
                    for row in rows:
                        if header == 1:
                            header = 0
                            ws.write(row[cols["header0"]] + "\n")
                            ws.write(row[cols["header1"]] + "\n")
                        ws.write(row[cols["fields"]] + "\n")

    def map_to_excel(self):
        # self.logger.info('Extract query : {0}'.format(self.config['DB']['market_price_export']))
        cursor = self.ds21_connection.cursor()
        for step in self.config['MAPPINGS']:
            if step == 'cols':
                cols = self.config['MAPPINGS'][step]
            else:
                sql = self.config['MAPPINGS'][step]
                sql = sql.replace('cols', cols)
                self.logger.info(sql)
                cursor.execute(sql)
                cursor.execute("commit")

        for xls in self.config['EXCEL']:
            file_name = get_file_name(self.config['FILES']['directory'], f"{xls}_<seqno>.xlsx")
            writer = pd.ExcelWriter(file_name, engine='openpyxl')
            self.logger.info(f"About to create excel file {file_name}")
            sql = self.config['EXCEL'][xls]
            self.logger.info(sql)
            df = pd.read_sql_query(sql, self.ds21_engine)
            df.to_excel(writer)
            self.logger.info('Excel file {0} created'.format(file_name))
            # Close the Pandas Excel writer and output the Excel file.
            writer.save()
            writer.close()

        cursor.close()

    def user_input(self):
        while True:
            try:
                exact_server = self.config['EXACT']['server']
                exact_server = input(f"Enter Exact server IP address: [{exact_server}] : ") or exact_server
                exact_server_ip = ipaddress.ip_address(exact_server)
                self.config['EXACT']['server'] = exact_server
                break
            except ValueError:
                print(f"Not a valid ip address : {exact_server}")
                continue

        while True:
            try:
                exact_username = self.config['EXACT']['username']
                exact_username = input(f"Enter Exact database username [{exact_username}] : ") or exact_username
                self.config['EXACT']['username'] = exact_username
                break
            except ValueError:
                print(f"Not a valid username : {exact_username}")
                continue

        while True:
            try:
                exact_password = self.config['EXACT']['password']
                exact_password = input(f"Password user {exact_username} in Exact database [{exact_password}] : ") \
                                 or exact_password
                self.config['EXACT']['password'] = exact_password
                break
            except ValueError:
                print(f"Not a valid password {exact_password}")
                continue

        while True:
            directory = self.config['FILES']['directory']
            directory = input(f"Directory Xcart import files [{directory}] : ") or directory
            if os.path.isdir(directory):
                self.config['FILES']['directory'] = directory
                break
            else:
                print(f"Geen directory: {directory}")

        while True:
            thumb_directory = self.config['PICTURES']['thumb_directory']
            thumb_directory = input(f"Directory files voor Xcart [{directory}] : ") or thumb_directory
            if os.path.isdir(directory):
                self.config['PICTURES']['thumb_directory'] = thumb_directory
                break
            else:
                print(f"Geen directory: {thumb_directory}")

        while True:
            plaatjes = self.config['PICTURES']['plaatjes']
            plaatjes = input(f"Ook fotos overhalen uit Exact (J/N) [{plaatjes}] ? ") or plaatjes
            if plaatjes in ('J', 'N'):
                self.config['PICTURES']['thumb_directory'] = thumb_directory
                break
            else:
                print(f"J of N, niet: {plaatjes}")

def get_engine(logger, config, db):
    db = db.upper()
    type_db = config[db]['type_db']
    server = config[db]['server']
    port = config[db]['port']
    database = config[db]['database']
    username = config[db]['username']
    password = config[db]['password']
    logger.info('Connecting to {0}'.format(database))
    if type_db == 'mysql':
        conn_string = 'mysql+pymysql://{0}:{1}@{2}/{3}'.format(username, password, server, database)
        engine = create_engine(conn_string, echo=False)
        connection = engine.raw_connection()
    elif type_db == 'mssql':
        quoted = f"DRIVER=FreeTDS;SERVER={server};PORT={port};DATABASE={database};UID={username};PWD={password}"
        conn_string = 'mssql+pyodbc:///?odbc_connect={}'.format(quoted)
        engine = create_engine(conn_string, echo=True)
        print(quoted)
        connection = pyodbc.connect(quoted)

    logger.info('Connection to {0} created'.format(db))

    return engine, connection


def get_file_name(directory, file_name):
    file_name = file_name.replace('<seqno>', time.strftime("%Y%m%d-%H%M%S"))
    file_name = os.path.join(directory, file_name)

    return file_name


def fields(cursor):
    """ Given a DB API 2.0 cursor object that has been executed, returns
    a dictionary that maps each field name to a column index; 0 and up. """
    results = {}
    column = 0
    for d in cursor.description:
        results[d[0]] = column
        column = column + 1

    return results


if __name__ == "__main__":
    winkel = Winkel()
    winkel.user_input()
    # winkel.ms_to_my()
    if winkel.config['PICTURES']['plaatjes'] == 'J':
        winkel.images()
    winkel.xcart_files()

    # winkel.map_to_excel()
    winkel.ds21_connection.close()
    winkel.ds21_engine.dispose()
    winkel.logger.info("Einde")
