import os
import sys
import csv
import pyodbc
import pypyodbc
import logging
import configparser
from time import sleep
from PySide2 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread
from PySide2.QtCore import QObject, Slot, Signal, QThread


class Pixmap(QtCore.QObject):
    def __init__(self, pix):
        super(Pixmap, self).__init__()

        self.pixmap_item = QtWidgets.QGraphicsPixmapItem(pix)
        self.pixmap_item.setCacheMode(QtWidgets.QGraphicsItem.DeviceCoordinateCache)

    def set_pos(self, pos):
        self.pixmap_item.setPos(pos)

    def get_pos(self):
        return self.pixmap_item.pos()

    pos = QtCore.Property(QtCore.QPointF, get_pos, set_pos)


def qt_message_handler(mode, context, message):
    pass


class View(QtWidgets.QGraphicsView):
    def resizeEvent(self, event):
        super(View, self).resizeEvent(event)
        self.fitInView(self.sceneRect(), QtCore.Qt.KeepAspectRatio)


class activator(QObject):
    sig = Signal()


names = list()
itemsSignals = list()
start = False
endSignal = activator()


@QtCore.Slot(list)
def setNames(Names):
    global names
    names = Names


class animateTiles(View):

    def __init__(self, parent=None):

        self.thread = databaseThread()
        self.thread.namesSignal.connect(setNames)
        self.thread.start()
        self.bgPix = QtGui.QPixmap('img/bg.jpg')
        self.scene = QtWidgets.QGraphicsScene(0, 0, 835, 470)

        self.timer = list()
        self.states = QtCore.QStateMachine()
        self.group = QtCore.QParallelAnimationGroup()

        self.items = []
        self.text = list()
        self.names = list()

        while (len(itemsSignals) != len(names) or len(itemsSignals) == 0 or len(names) == 0):
            continue

        self.names = names
        for i, name in enumerate(self.names):
            self.kineticPix =  QtGui.QPixmap('img/' + name + '.png')
            # for i in range(len(names)):
            self.item = Pixmap(self.kineticPix)
            self.item.pixmap_item.setOffset(self.kineticPix.width() * 0.75 + 82,
                                       -self.kineticPix.height() * 0.25)
            self.item.pixmap_item.setZValue(len(names)-i)
            self.items.append(self.item)
            self.scene.addItem(self.item.pixmap_item)
            self.text.append(self.scene.addText("Loading " + name + '...'))
            self.text[i].setScale(1.25)

        self.endPix = QtGui.QPixmap('img/endScene.png')
        self.endItem = Pixmap(self.endPix)
        self.endItem.pixmap_item.setOffset(95,0)
        self.endItem.pixmap_item.setZValue(1)
        self.scene.addItem(self.endItem.pixmap_item)

        self.tiledStates = list()
        # States.
        self.rootState = QtCore.QState()
        for i in range(0, len(names)):
            self.tiledStates.append(QtCore.QState(self.rootState))
        self.centeredState = QtCore.QState(self.rootState)
        self.endState = QtCore.QState(self.rootState)

        # Values.
        for i, item in enumerate(self.items):
            # Tiled.
            self.tiledStates[i].assignProperty(item, 'pos',
                                                           QtCore.QPointF((i % 5 )* self.kineticPix.width()*1.2 - self.kineticPix.width() / 2,
                                                                          (i // 5) * self.kineticPix.height()*1.2 + self.kineticPix.height() / 2))
            for j in range(0, len(self.names)):
                self.tiledStates[i].assignProperty(self.text[j], 'pos',
                                              QtCore.QPointF(-1000, 440))
            self.tiledStates[i].assignProperty(self.text[i], 'pos',
                                          QtCore.QPointF(10, 440))
            # Centered.
            self.centeredState.assignProperty(item, 'pos',
                                         QtCore.QPointF(312 - (self.kineticPix.width() * len(self.items) * 0.5 * 0.3) - (self.kineticPix.width() * 0.5) + i * self.kineticPix.width() * 0.3, -50 - self.kineticPix.height() * 0.5))
            self.centeredState.assignProperty(self.endItem, 'pos',
                                    QtCore.QPointF(0, 1000))
            self.centeredState.assignProperty(self.text[i], 'pos',
                                               QtCore.QPointF(10, 1000))
            # End state.
            self.endState.assignProperty(item, 'pos',
                                         QtCore.QPointF(312 - self.kineticPix.width() * 1.25, -1000))
            self.endState.assignProperty(self.text[i], 'pos',
                                          QtCore.QPointF(10, 1000))
            self.endState.assignProperty(self.endItem, 'pos',
                                    QtCore.QPointF(0, 145))

        # Ui.
        self.view = View(self.scene)
        self.view.setWindowTitle("Altium Library Converter")
        self.view.setViewportUpdateMode(QtWidgets.QGraphicsView.BoundingRectViewportUpdate)
        self.view.setBackgroundBrush(QtGui.QBrush(self.bgPix))
        self.view.setCacheMode(QtWidgets.QGraphicsView.CacheBackground)
        self.view.setRenderHints(
            QtGui.QPainter.Antialiasing | QtGui.QPainter.SmoothPixmapTransform)
        self.view.show()

    def run(self):
        self.states.addState(self.rootState)
        self.states.setInitialState(self.rootState)
        self.rootState.setInitialState(self.centeredState)

        for i, item in enumerate(self.items):
            anim = QtCore.QPropertyAnimation(item, b'pos')
            anim.setDuration(750)
            anim.setEasingCurve(QtCore.QEasingCurve.InOutBack)
            self.group.addAnimation(anim)

        self.anim = QtCore.QPropertyAnimation(self.endItem, b'pos')
        self.anim.setDuration(750)
        self.anim.setEasingCurve(QtCore.QEasingCurve.InOutBack)
        self.group.addAnimation(self.anim)

        self.trans = self.rootState.addTransition(endSignal.sig, self.endState)
        self.trans.addAnimation(self.group)

        for i in range(0, len(self.names)):
            self.trans = self.rootState.addTransition(itemsSignals[i].sig, self.tiledStates[i])
            self.trans.addAnimation(self.group)

        self.states.start()
        global start
        start = True


class databaseThread(QThread):

    namesSignal = Signal(list)
    def __init__(self, parent=None):
        QThread.__init__(self, parent)
        self.exiting = False
        self.names = list()

    def run(self):
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
            config['GENERAL'] = {'ALTIUM_LIB_PATH': 'default',
                                 'ODBC_DRIVER': 'Microsoft Access Driver (*.mdb)',
                                 'OPEN_ALTIUM': 'false',
                                 'ALTIUM_PATH': 'C:\Program Files\Altium\AD19\X2.EXE',
                                 'GIT_CHECKOUT.': 'false',
                                 'GIT_PULL': 'false'}

            with open('config.ini', 'w') as configfile:
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
                                 'formlas_conflict': 'Device,Family,Value,Color,Manufacturer 1 Part Number'}

            with open('settings.ini', 'w') as configfile:
                settings.write(configfile)

        config = configparser.RawConfigParser()
        config.optionxform = str

        settings = configparser.RawConfigParser()
        settings.optionxform = str

        try:
            if os.stat("config.ini").st_size == 0:
                createConfig()
        except PermissionError:
            pass
        except FileNotFoundError:
            createConfig()

        config.read('config.ini')

        try:
            if os.stat("settings.ini").st_size == 0:
                createSettings()
        except PermissionError:
            pass
        except FileNotFoundError:
            createSettings()

        settings.read('settings.ini')
        if config['GENERAL']['ALTIUM_LIB_PATH'] == 'default':
            libPath = os.getcwd()
        else:
            libPath = config['GENERAL']['ALTIUM_LIB_PATH']

        if config.getboolean('GENERAL', 'GIT_CHECKOUT.'):
            os.popen("git checkout .")

        if config.getboolean('GENERAL', 'GIT_PULL'):
            os.popen("git pull")

        def constructDB(name, tables):
            return {
                'name': name,
                'tables': tables
            }
        noerror = 0
        line = list()
        databases = list()
        for key in settings['STRUCTURE']:
            names.append(key)
            databases.append(constructDB(key, settings['STRUCTURE'][key].split(",")))

        for i in range(0, len(names)):
            itemsSignals.append(activator())

        while not start:
            continue
        sleep(0.1)
        for i, database in enumerate(databases):
            itemsSignals[i].sig.emit()
            for table in database['tables']:

                try:
                    # TEXT FILE CLEAN
                    CSVfilepath = "CSV/" + table + ".csv"
                    with open(os.path.join(libPath, CSVfilepath), 'r') as reader, open(os.path.join(dirname, 'CSV/temp.csv'), 'w') as tempwriter, open(os.path.join(dirname, 'CSV/Clean.csv'), 'w') as writer:
                        read_csv = csv.reader(reader, delimiter=','); tempwrite_csv = csv.writer(tempwriter, lineterminator='\n'); write_csv = csv.writer(writer, lineterminator='\n')
                        for ReadedLine in read_csv:
                            try:
                                if len(ReadedLine[1]) > 0:
                                    del line[:]
                                    for word in ReadedLine:
                                        if settings.getboolean('EXCEL', 'set_formulas_as_text'):
                                            for conflict in settings['EXCEL']['formlas_conflict'].split(","):
                                                if word is None:
                                                    continue
                                                word = word.replace("'=" + conflict + "'", "=" + conflict)
                                        line.append(word)
                                    write_csv.writerow(line)

                            except IndexError:
                                logging.error(f"EMPTY CSV: {table}.csv")
                                #print(f"     EMPTY CSV: {table}.csv")
                                noerror = 0
                                '''
                            except AttributeError:
                                print(f"     atribute")
                                noerror = 0
                                '''
                            else:
                                noerror = 1
                    if noerror:
                        reader.close()
                        tempwriter.close()
                        writer.close()
                        #shutil.move('CSV/temp.csv', CSVfilepath)

                        MDBfilepath = "DataBase/" + database['name'] + ".mdb"
                        access_path = os.path.join(libPath, MDBfilepath)

                        if not os.path.exists(os.path.join(libPath, "DataBase/")):
                            os.makedirs(os.path.join(libPath, "DataBase/"))

                        if not os.path.exists(access_path):
                            pypyodbc.win_create_mdb(access_path)  # CREATE MDB IF NOT EXIST
                            logging.info(f"CREATING MDB: {database['name']}.mdb")
                            #print(f"CREATING MDB: {database['name']}.mdb")

                        # DATABASE CONNECTION
                        con = pyodbc.connect("DRIVER={" + config['GENERAL']['ODBC_DRIVER'] + "};" + \
                                            "DBQ=" + access_path + ";")
                        cur = con.cursor()

                        if not cur.tables(table= table, tableType='TABLE').fetchone():
                            strSQL = "CREATE TABLE " + table    #CREATE TABLE IF NOT EXIST
                            logging.info(f"CREATING TABLE: {table} in {database['name']}.mdb")
                            #print(f"CREATING TABLE: {table} in {database['name']}.mdb")
                            cur.execute(strSQL)

                        # RUN QUERY
                        strSQL = "DROP TABLE " + table
                        logging.info(f"CONVERTING CSV: {table}.csv")
                        #print(f"CONVERTING CSV: {table}.csv")

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
                    #print(f" NOT FOUND CSV: {table}.csv")
                except pyodbc.InterfaceError:
                    noerror = 0
                    logging.error("ODBC DRIVER ERROR")
                    #print("ODBC DRIVER ERROR")
        #print("DONE!")
        endSignal.sig.emit()

        if config.getboolean('GENERAL', 'OPEN_ALTIUM'):
            os.startfile(config['GENERAL']['ALTIUM_PATH'])

        sleep(0.700)
        sys.exit(0)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    QtCore.qInstallMessageHandler(qt_message_handler)
    window = animateTiles()
    window.run()
    sys.exit(app.exec_())