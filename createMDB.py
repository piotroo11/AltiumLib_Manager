import os
import csv
import pyodbc
import pypyodbc
import logging
import configparser
import datetime
import shutil
"""
LIST OF DATABASE STRUCTURE
Antennas And RF Components -> Antennas
Capacitors -> Capacitors_SMD, Capacitors_THD
Connectors -> Connectors_SMD, Connectors_THD
ICs And Semiconductors -> ADCs, Amplifiers, Crystals_and_Oscillators, DACs, Diodes, LEDs, Memories, Microcontrollers, Regulators, Transistors
Inductors -> Inductors_SMD, Inductors_THD
Mechanical And PCB Features -> Mechanical, PCB_Features
Relays And Switches -> Relays_SMD, Relays_THD, Switches_SMD, Switches_THD
Resistors -> Resistors_SMD, Thermistors_SMD, Thermistors_THD
"""

print("       /\   | | |__   __|_   _| |  | |  \/  |     / ____|/ ____|  __ \|_   _|  __ \__   __|               ")
print("      /  \  | |    | |    | | | |  | | \  / |    | (___ | |    | |__) | | | | |__) | | |                  ")
print("     / /\ \ | |    | |    | | | |  | | |\/| |     \___ \| |    |  _  /  | | |  ___/  | |                  ")
print("    / ____ \| |____| |   _| |_| |__| | |  | |     ____) | |____| | \ \ _| |_| |      | |                  ")
print("   /_/    \_\______|_|  |_____|\____/|_|  |_|    |_____/ \_____|_|  \_\_____|_|      |_|  by Piotr Czapnik")

dirname = os.getcwd()
try:
    if os.stat("log.txt").st_size != 0:
        os.remove(os.path.join(dirname, 'log.txt'))  # CLEAN OLD LOG
except PermissionError:
    pass
except FileNotFoundError:
    pass

logging.basicConfig(filename="log.txt", level=logging.INFO)


def createConfig():
    config = configparser.RawConfigParser()
    config.optionxform = str
    config['GENERAL'] = {'ODBC_DRIVER': 'Microsoft Access Driver (*.mdb)',
                         'OPEN_ALTIUM': 'false',
                         'ALTIUM_PATH': 'C:\Program Files\Altium\AD19',
                         'GIT_CHECKOUT.': 'false',
                         'GIT_PULL': 'false'}

    with open('config2.ini', 'w') as configfile:
        config.write(configfile)

def createSettings():
    settings = configparser.RawConfigParser()
    settings.optionxform = str

    settings['STRUCTURE'] = {'Antennas And RF Components': 'Antennas',
                         'Capacitors': 'Capacitors_SMD,Capacitors_THD',
                         'Connectors': 'Connectors_SMD,Connectors_THD',
                         'ICs And Semiconductors': 'ADCs,Amplifiers,Amplifiers,Crystals_and_Oscillators,DACs,Diodes,LEDs,Memories,Microcontrollers,Regulators,Transistors',
                         'Inductors': 'Inductors_SMD,Inductors_THD',
                         'Mechanical And PCB Features': 'Mechanical,PCB_Features',
                         'Relays And Switches': 'Relays_SMD,Relays_THD,Switches_SMD,Switches_THD',
                         'Resistors': 'Resistors_SMD,Thermistors_SMD,Thermistors_THD'}

    settings['EXCEL'] = {'set_formulas_as_text': 'true',
                         'formlas_conflict': 'Device,Family,Value,Color,Manufacturer 1 Part Number',
                         'set_date_as_text': 'true'}

    with open('settings.ini', 'w') as configfile:
        settings.write(configfile)


config = configparser.RawConfigParser()
config.optionxform = str

settings = configparser.RawConfigParser()
settings.optionxform = str

try:
    if os.stat("config2.ini").st_size == 0:
        createConfig()
except PermissionError:
    pass
except FileNotFoundError:
    createConfig()

config.read('config2.ini')

try:
    if os.stat("settings.ini").st_size == 0:
        createSettings()
except PermissionError:
    pass
except FileNotFoundError:
    createSettings()

settings.read('settings.ini')

if config.getboolean('GENERAL', 'GIT_CHECKOUT.'):
    os.popen("git checkout .")
if config.getboolean('GENERAL', 'GIT_PULL'):
    os.popen("git pull")


def constructDB(name, tables):
    return {
        'name': name,
        'tables': tables
    }
'''
line = list()
tempword = str()
templine = list()
conflict = str()
def repairData(word):
    global line
    global templine
    global tempword
    global conflict
    if word is None:
        return continue
    word = str(word)
    word = word.replace("'=" + conflict + "'", "=" + conflict)
    tempword = word
    try:
        word = word.replace("'" + datetime.datetime.strptime(word, "%d/%m/%y").strftime("%d/%m/%y") + "'",
                                datetime.datetime.strptime(word, "%d/%m/%y").strftime("%d.%m.%y"))
        tempword = word.replace(datetime.datetime.strptime(word, "%d/%m/%y").strftime("%d/%m/%y"),
                                    "'" + datetime.datetime.strptime(word, "%d/%m/%y").strftime("%d.%m.%y") + "'")
        print('1')
    except ValueError as err:
        try:
            word = word.replace("'" + datetime.datetime.strptime(word, "%d.%m.%y").strftime("%d.%m.%y") + "'",
                                    datetime.datetime.strptime(word, "%d.%m.%y").strftime("%d.%m.%y"))
            tempword = word.replace(datetime.datetime.strptime(word, "%d.%m.%y").strftime("%d.%m.%y"),
                                    "'" + datetime.datetime.strptime(word, "%d.%m.%y").strftime("%d.%m.%y" + "'"))
            print('2')
        except ValueError as err:
            try:

                word = word.replace("'" + datetime.datetime.strptime(word, "%d-%m-%y").strftime("%d-%m-%y") + "'",
                                        datetime.datetime.strptime(word, "%d-%m-%y").strftime("%d.%m.%y"))
                tempword = word.replace(datetime.datetime.strptime(word, "%d-%m-%y").strftime("%d-%m-%y"),
                                     "'" + datetime.datetime.strptime(word, "%d-%m-%y").strftime("%d.%m.%y") + "'")
                #print(word)
                #print(tempword)
            except ValueError as err:
                pass
    print(line)
    templine.append(tempword)
    #print(templine)
'''

noerror = 0
templine = list()
line = list()
word = str()
databases = list()
for key in settings['STRUCTURE']:
     databases.append(constructDB(key, settings['STRUCTURE'][key].split(",")))

for database in databases:
    for table in database['tables']:

        try:
            # TEXT FILE CLEAN
            CSVfilepath = "CSV/" + table + ".csv"
            with open(os.path.join(dirname, CSVfilepath), 'r') as reader, open(os.path.join(dirname, 'CSV/temp.csv'), 'w') as tempwriter, open(os.path.join(dirname, 'CSV/Clean.csv'), 'w') as writer:
                read_csv = csv.reader(reader, delimiter=','); tempwrite_csv = csv.writer(tempwriter, lineterminator='\n'); write_csv = csv.writer(writer, lineterminator='\n')
                for ReadedLine in read_csv:
                    try:
                        if len(ReadedLine[1]) > 0:
                            for word in ReadedLine:
                                if settings.getboolean('EXCEL', 'set_formulas_as_text'):
                                    for conflict in settings['EXCEL']['formlas_conflict'].split(","):
                                        del line[:]

                                            if word is None:
                                                print('huj')
                                                continue
                                            #word = str(word)
                                            word = word.replace("'=" + conflict + "'", "=" + conflict)
                                            '''
                                            tempword = word
                                            try:
                                                word = word.replace(
                                                    "'" + datetime.datetime.strptime(word, "%d/%m/%y").strftime(
                                                        "%d/%m/%y") + "'",
                                                    datetime.datetime.strptime(word, "%d/%m/%y").strftime("%d.%m.%y"))
                                                tempword = word.replace(
                                                    datetime.datetime.strptime(word, "%d/%m/%y").strftime("%d/%m/%y"),
                                                    "'" + datetime.datetime.strptime(word, "%d/%m/%y").strftime(
                                                        "%d.%m.%y") + "'")
                                                print('1')
                                            except ValueError as err:
                                                try:
                                                    word = word.replace(
                                                        "'" + datetime.datetime.strptime(word, "%d.%m.%y").strftime(
                                                            "%d.%m.%y") + "'",
                                                        datetime.datetime.strptime(word, "%d.%m.%y").strftime("%d.%m.%y"))
                                                    tempword = word.replace(
                                                        datetime.datetime.strptime(word, "%d.%m.%y").strftime("%d.%m.%y"),
                                                        "'" + datetime.datetime.strptime(word, "%d.%m.%y").strftime(
                                                            "%d.%m.%y" + "'"))
                                                    print('2')
                                                except ValueError as err:
                                                    try:
    
                                                        word = word.replace(
                                                            "'" + datetime.datetime.strptime(word, "%d-%m-%y").strftime(
                                                                "%d-%m-%y") + "'",
                                                            datetime.datetime.strptime(word, "%d-%m-%y").strftime(
                                                                "%d.%m.%y"))
                                                        tempword = word.replace(
                                                            datetime.datetime.strptime(word, "%d-%m-%y").strftime(
                                                                "%d-%m-%y"),
                                                            "'" + datetime.datetime.strptime(word, "%d-%m-%y").strftime(
                                                                "%d.%m.%y") + "'")
                                                        print('3')
                                                        # print(word)
                                                        # print(tempword)
                                                    except ValueError as err:
                                                        pass
                                            print(line)
                                            templine.append(tempword)
                                            '''
                                            line.append(word)
                                write_csv.writerow(line)
                                #print(line)
                                '''
                            if settings.getboolean('EXCEL', 'set_date_as_text'):
                                tempwrite_csv.writerow(templine)
                                #print(templine)
                                '''
                    except IndexError:
                        logging.error(f"EMPTY CSV: {table}.csv")
                        print(f"     EMPTY CSV: {table}.csv")
                        noerror = 0
                        '''
                    except AttributeError:
                        print(f"     atribute")
                        noerror = 0
                        '''
                    else:
                        noerror = 1
            if noerror:
                #reader.close()
                #tempwriter.close()
                #writer.close()
                #shutil.move('CSV/temp.csv', CSVfilepath)

                MDBfilepath = "DataBase/" + database['name'] + ".mdb"
                access_path = os.path.join(dirname, MDBfilepath)

                if not os.path.exists(os.path.join(dirname, "DataBase/")):
                    os.makedirs(os.path.join(dirname, "DataBase/"))

                if not os.path.exists(access_path):
                    pypyodbc.win_create_mdb(access_path)  # CREATE MDB IF NOT EXIST
                    logging.info(f"CREATING MDB: {database['name']}.mdb")
                    print(f"CREATING MDB: {database['name']}.mdb")

                # DATABASE CONNECTION
                con = pyodbc.connect("DRIVER={" + config['GENERAL']['ODBC_DRIVER'] + "};" + \
                                    "DBQ=" + access_path + ";")
                cur = con.cursor()

                if not cur.tables(table= table, tableType='TABLE').fetchone():
                    strSQL = "CREATE TABLE " + table    #CREATE TABLE IF NOT EXIST
                    logging.info(f"CREATING TABLE: {table} in {database['name']}.mdb")
                    print(f"CREATING TABLE: {table} in {database['name']}.mdb")
                    cur.execute(strSQL)

                # RUN QUERY
                strSQL = "DROP TABLE " + table
                logging.info(f"CONVERTING CSV: {table}.csv")
                print(f"CONVERTING CSV: {table}.csv")

                cur.execute(strSQL)

                strSQL = "SELECT * INTO " + table + " FROM [text;HDR=Yes;FMT=Delimited(,);" + \
                        "Database=" + os.path.join(dirname, 'CSV') + "].Clean.csv"
                cur = con.cursor()
                cur.execute(strSQL)
                con.commit()

                con.close()                            # CLOSE CONNECTION
                os.remove(os.path.join(dirname, 'CSV/Clean.csv'))       # DELETE CLEAN TEMP
        except FileNotFoundError:
            noerror = 0
            try:
                os.remove(os.path.join(dirname, 'CSV/Clean.csv'))  # DELETE CLEAN TEMP
            except PermissionError:
                pass
            except FileNotFoundError:
                pass
            logging.error(f"NOT FOUND CSV: {table}.csv")
            print(f" NOT FOUND CSV: {table}.csv")
        except pyodbc.InterfaceError:
            noerror = 0
            logging.error("ODBC DRIVER ERROR")
            print("ODBC DRIVER ERROR")
print("DONE!")

if config.getboolean('GENERAL', 'OPEN_ALTIUM'):
    os.startfile(config['GENERAL']['ALTIUM_PATH'] + "\X2.EXE")

