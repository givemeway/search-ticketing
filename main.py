
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, \
    QTableWidgetItem, QMessageBox, QLabel, QHeaderView
from PyQt5.QtCore import QObject, pyqtSignal, QThread, pyqtSlot, QTimer
from PyQt5.QtGui import QMovie, QIcon
from PyQt5 import QtGui
import sys
import os
import csv
from parse import error
from math import ceil
from parse import get_mbox, process_emls
from findAndLoad import loadPrevious, search_ticket
from new_ui_updated import Ui_TicketingSearchTool
from login import *
from searchTicket import ticketing, get_path
from escalation_tracker import escalation_worker, agent_tracker_worker
from copy import deepcopy
# import webbrowser

# implement
# https://stackoverflow.com/questions/12009134/adding-widgets-to-qtablewidget-pyqt
# https://learndataanalysis.org/create-hyperlinks-pyqt5-tutorial/


class HyperlinkLabel(QLabel):
    def __init__(self, parent):
        self.parent = parent
        super().__init__(self.parent)
        self.setOpenExternalLinks(True)
        # self.setParent(parent)


class LoadWorker(QObject):
    finished = pyqtSignal()
    email_not_found = pyqtSignal(list)
    subject_not_found = pyqtSignal(list)

    def __init__(self, id):
        super().__init__()
        self.id = id

    def run(self):
        _email_not_found = self.email_not_found
        _subject_not_found = self.subject_not_found
        loadPrevious(self.id, _email_not_found, _subject_not_found)
        self.finished.emit()


class QuerySearchWorker(QObject):
    finished = pyqtSignal()
    result = pyqtSignal(object)

    def __init__(self, sess, query):
        super().__init__()
        self.sess = sess
        self.query = query

    def run(self):
        _result = self.result
        search_ticket(self.sess, self.query, _result)
        self.finished.emit()


class LoginWorker(QObject):
    finished = pyqtSignal()
    session = pyqtSignal(object)
    result = pyqtSignal(int)
    error = pyqtSignal(tuple)

    def __init__(self, username, password):
        super().__init__()
        self.username = username
        self.password = password

    def run(self):
        _session = self.session
        status = ticket_session(self.username, self.password, _session)
        if status == 200:
            self.result.emit(200)
        elif status == 403:
            self.result.emit(403)
        elif status == 422:
            self.result.emit(422)
        else:
            self.error.emit(status)
        self.finished.emit()


class ParserWorker(QObject):
    finished = pyqtSignal()
    loading = pyqtSignal(tuple)
    result = pyqtSignal(list)
    error = pyqtSignal(tuple)
    progress = pyqtSignal(list)
    progress_bar = pyqtSignal(tuple)

    def __init__(self, _dir, module):
        super().__init__()
        self.module = module
        self.dir = _dir

    def run(self):
        _progress = self.progress
        _loading = self.loading
        _progress_bar = self.progress_bar
        _error = self.error
        if self.module == 'mbox':
            self.emails = get_mbox(self.dir, _progress,
                                   _loading, _progress_bar)
            if isinstance(self.emails, list):
                self.result.emit(self.emails)
            else:
                self.error.emit(('MBox File(s) not found', 'Please try again'))
        elif self.module == 'eml':
            self.emails = process_emls(
                self.dir, _progress, _loading, _progress_bar, _error)
            if isinstance(self.emails, list):
                self.result.emit(self.emails)
            else:
                self.error.emit(('Eml File(s) not found', 'Please try again'))
        self.finished.emit()


class EscalationWorker(QObject):
    finished = pyqtSignal()
    pBar = pyqtSignal(int)
    error = pyqtSignal(tuple)
    warning = pyqtSignal(object)
    processing = pyqtSignal(tuple)

    def __init__(self, tickets, session, tracker):
        super().__init__()
        self.tickets = tickets
        self.session = session
        self.tracker = tracker

    def run(self):
        _dict = {'tickets': self.tickets,
                 'session': self.session,
                 'warning': self.warning,
                 'processing': self.processing,
                 'pBar': self.pBar,
                 'error': self.error,
                 'total': len(self.tickets),
                 }
        if self.tracker == "AGENT":
            agent_tracker_worker(_dict)
        elif self.tracker == "ESCALATION":
            escalation_worker(_dict)
        self.finished.emit()


class SearchWorker(QObject):
    finished = pyqtSignal()
    pBar = pyqtSignal(int)
    error = pyqtSignal(tuple)
    sessionError = pyqtSignal(bool)
    expiredAt = pyqtSignal(object)
    warning = pyqtSignal(bool)
    processing = pyqtSignal(tuple)
    email_not_found = pyqtSignal(list)
    subject_not_found = pyqtSignal(list)

    def __init__(self, emails, session, startFrom, sessionExpired=False):
        super().__init__()
        self.emails = emails
        self._sess = session
        self.startFrom = startFrom
        self.sessionExpired = sessionExpired

    def run(self):
        _dict = {'emails': self.emails,
                 'session': self._sess,
                 'warning': self.warning,
                 'processing': self.processing,
                 'pBar': self.pBar,
                 'emailNotFound': self.email_not_found,
                 'subjectNotFound': self.subject_not_found,
                 'error': self.error,
                 'emailCount': len(self.emails),
                 'sessionError': self.sessionError,
                 'expiredAt': self.expiredAt,
                 'startFrom': self.startFrom, }
        ticketing(_dict)
        self.finished.emit()


class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_TicketingSearchTool()
        self.ui.setupUi(self)
        self.cols = ["SI NO", "E-mail", "Subject", "Date Sent"]
        self.cols_sub_table = ["SI NO", "E-mail",
                               "Match", "Subject", "Date Sent"]
        self.ui.loading.setText("")
        self.ui.btn_find.setDisabled(True)
        self.ui.btn_parser.setDisabled(True)
        self.ui.btn_transferred.setDisabled(True)
        # file dialogue
        self.ui.toolButton.clicked.connect(self.mboxEmlDirPath)
        # connect the buttons to their respective actions
        self.ui.btn_login.clicked.connect(self.login)
        self.ui.pushButton_4.clicked.connect(self.loadComboBox)
        self.ui.comboBox.currentIndexChanged.connect(self.comboxSelected)
        self.ui.pushButton_3.clicked.connect(self.loadPreviousResults)
        self.ui.pushButton_2.clicked.connect(self.querySearch)
        self.ui.pushButton.clicked.connect(self.startSearch)
        self.ui.btn_parse_option.clicked.connect(self.parse)
        self.ui.btn_parse_clear.clicked.connect(
            lambda: self.clearTable(self.ui.tableWidget))
        # self.ui.tableWidget_6.itemDoubleClicked.connect( self.OpenLink)
        # connecting side buttons with the widget page
        self.ui.btn_parser.clicked.connect(
            lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.parser_page))
        self.ui.btn_find.clicked.connect(
            lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.find_page))
        self.ui.btn_transferred.clicked.connect(
            lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.transferred_page))
        # show and hide Search Underway
        self.ui.btn_parser.clicked.connect(
            lambda: self.ui.frame_6.setHidden(True))
        self.ui.btn_find.clicked.connect(self.searchUnderway)
        self.ui.btn_transferred.clicked.connect(self.searchUnderway)
        # show the parser search screen
        self.ui.pushButton_5.clicked.connect(
            lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.parser_page)
        )
        self.ui.pushButton_5.clicked.connect(
            lambda: self.ui.frame_6.setHidden(True))
        # auto login if the app is previously logged in

        self.ui.pushButton_7.clicked.connect(self.enableAutoSearch)

        # escalated tracker
        # ============================================================ #
        self.ui.tracker_csv_toolButton.clicked.connect(self.selectCSV)
        self.ui.tracker_pushButton.setHidden(True)
        self.ui.tracker_pushButton.clicked.connect(self.escalationSearch)

        # ============================================================ #
        self.isSessionExpired = False
        self.auto_login()
        self.comboIndex = None
        self.font = QtGui.QFont()
        self.font.setFamily("Verdana")
        self.font.setPointSize(9)
        self.font.setBold(True)
        self.font.setWeight(75)
        # set search to false
        self.isSearch = False
        # set inital mbox or eml path
        self.EmlMBoxPath = None
        self.startFrom = {"idx": 0, "email_idx": 0,
                          "subject_idx": 0, "err_idx": 0}
        self.failedAt = {"idx": 0, "email_idx": 0,
                         "subject_idx": 0, "err_idx": 0}
        self.expiredAtCount = 0
        self.username = None
        self.password = None
        # setup context menu in Table
        # https://stackoverflow.com/questions/50768366/installeventfilter-in-pyqt5
        # https://stackoverflow.com/questions/65371143/create-a-context-menu-with-pyqt5/65371906#65371906
        # https://stackoverflow.com/questions/51619186/pyqt5-qtablewidget-right-click-on-cell-wont-spawn-qmenu
        self.ui.tableWidget.customContextMenuRequested.connect(
            self.ui.tableWidget.generateMenu)
        self.ui.tableWidget_2.customContextMenuRequested.connect(
            self.ui.tableWidget_2.generateMenu)
        self.ui.tableWidget_3.customContextMenuRequested.connect(
            self.ui.tableWidget_3.generateMenu)
        self.ui.tableWidget_4.customContextMenuRequested.connect(
            self.ui.tableWidget_4.generateMenu)
        self.ui.tableWidget_5.customContextMenuRequested.connect(
            self.ui.tableWidget_5.generateMenu)
        self.ui.tableWidget_6.customContextMenuRequested.connect(
            self.ui.tableWidget_6.generateMenu)

        self.ui.tableWidget_5.itemSelectionChanged.connect(
            lambda: self.autoSearch(self.ui.tableWidget_5))
        self.ui.tableWidget_6.itemSelectionChanged.connect(
            lambda: self.autoSearch(self.ui.tableWidget_6))

        self.isSearching = False
        self.isAutoSearch = False

        self.green_clip_icon = QIcon(get_path("icons/paper-clip.png"))
        self.red_clip_icon = QIcon(get_path("icons/attachment.png"))

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        close = QMessageBox.question(
            self,
            'QUIT',
            'Are you sure you want to quit?',
            QMessageBox.Yes | QMessageBox.No)
        if close == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def disableTabs(self, hide=True):
        self.ui.btn_find.setDisabled(hide)
        self.ui.btn_parser.setDisabled(hide)
        self.ui.btn_transferred.setDisabled(hide)

    def displaymsg(self, widget, text, style="color: rgb(255, 8, 12);\nfont: 75 10pt \"Verdana\";"):
        widget.setText(text)
        widget.setStyleSheet(style)

    def selectCSV(self):
        msg = "Select CSV File with Tickets"
        filter = "CSV Files (*.csv)"
        try:
            if os.path.exists("./logs/csvPath.log"):
                with open("./logs/csvPath.log") as file:
                    path = file.read()
                    if len(path) > 0:
                        prevCSVPath = path
                    else:
                        prevCSVPath = ""
            else:
                with open("./logs/csvPath.log", "w") as file:
                    file.write("")
                prevCSVPath = ""
        except Exception as e:
            error.exception(e)

        self.csvPath, _ = QFileDialog.getOpenFileName(
            self, msg, prevCSVPath, filter)
        if len(self.csvPath) > 0:
            self.ui.tracker_pushButton.setHidden(False)
            try:
                with open("./logs/csvPath.log", 'w') as file:
                    file.write(os.path.dirname(self.csvPath))
            except Exception as e:
                error.exception(e)

    def escalationSearch(self):
        tracker = None
        if self.ui.agent_radioButton.isChecked():
            tracker = "AGENT"
        elif self.ui.escalation_radioButton.isChecked():
            tracker = "ESCALATION"
        else:
            return
        self.time = 0
        self.timer = QTimer()
        self.timer.start(1000)
        self.timer.timeout.connect(self.Counter)
        self.tickets = []
        self.ui.escalation_progressBar.setValue(0)
        self.tickets = self.extract_tickets(self.csvPath)
        self.text_gif = QMovie(get_path('gifs/text_fading.gif'))
        self.ui.escalation_groupLabel.setMovie(self.text_gif)
        self.startAnimation(self.text_gif)

        self.escalationworker = EscalationWorker(
            self.tickets, self.session, tracker)
        self.escalationThread = QThread()
        self.escalationworker.moveToThread(self.escalationThread)
        self.escalationThread.started.connect(self.escalationworker.run)
        self.escalationworker.finished.connect(self.escalationThread.quit)
        self.escalationworker.finished.connect(
            self.escalationworker.deleteLater)
        self.escalationworker.finished.connect(lambda: self.showDialog(
            "Escalation Tracker", QMessageBox.Information, "Escalation Tracker Complete!" if tracker == "ESCALATION" else "Agent Tracker Complete!"))
        self.escalationworker.finished.connect(
            lambda: self.ui.tracker_pushButton.setEnabled(True))
        self.escalationworker.finished.connect(
            lambda: self.stopAnimation(self.text_gif))
        self.escalationworker.finished.connect(lambda: self.displaymsg(
            self.ui.escalation_groupLabel, text="Search Complete!", style="color: rgb(0, 98, 163);"))
        self.escalationworker.error.connect(self.escalationError)
        self.escalationworker.processing.connect(self.escalationProgress)
        self.escalationworker.pBar.connect(self.escalationProgressBar)
        self.escalationThread.finished.connect(
            self.escalationThread.deleteLater)
        self.escalationworker.warning.connect(self.escalationWarning)
        self.escalationThread.start()

        self.ui.escalation_processing_groupBox.setHidden(False)
        self.ui.tracker_pushButton.setEnabled(False)

    @pyqtSlot(int)
    def escalationProgressBar(self, per):
        self.ui.escalation_progressBar.setValue(per)

    @pyqtSlot(tuple)
    def escalationError(self, error):
        self.showDialog('Search Error', QMessageBox.Critical, f'''{error}''')

    @pyqtSlot(object)
    def escalationWarning(self, warning):
        self.showDialog('Warning Error', QMessageBox.Critical,
                        f'''{warning}''')

    @pyqtSlot(tuple)
    def escalationProgress(self, processed):
        self.ui.escalation_label_1.setText(f"Processing     : {processed[0]} ")
        self.ui.escalation_label_2.setText(
            f"Processed      : [ {processed[1]} of {len(self.tickets)} ]")
        self.ui.escalation_label_3.setText(
            f"Time Elapsed   : {self.convertSecsToHHMMSS(self.time)}")

    def extract_tickets(self, csvFilePath):
        img_obj = []
        columns = ['ticket']
        try:
            with open(csvFilePath, 'r', encoding='utf-8') as file:
                reader = csv.reader(file, delimiter=',')
                for row in reader:
                    if row != columns:
                        img_obj.append(row[0])
            return img_obj
        except Exception as e:
            error.exception(e)
            return e

    def mboxEmlDirPath(self):
        try:
            if os.path.exists("./logs/prevDir.log"):
                with open("./logs/prevDir.log") as file:
                    path = file.read()
                    if len(path) > 0:
                        prevMBoxPath = path
                    else:
                        prevMBoxPath = ""
            else:
                with open("./logs/prevDir.log", "w") as file:
                    file.write("")
                prevMBoxPath = ""
        except Exception as e:
            error.exception(e)

        self.EmlMBoxPath = QFileDialog.getExistingDirectory(
            self, "Choose MBox/Eml folder", prevMBoxPath)
        if len(self.EmlMBoxPath) > 0:
            self.ui.btn_parse_option.setHidden(False)
            try:
                with open("./logs/prevDir.log", 'w') as file:
                    file.write(self.EmlMBoxPath)
            except Exception as e:
                error.exception(e)

    def showDialog(self, title, displayIcon, displayText):
        msg = QMessageBox()
        msg.setIcon(displayIcon)
        msg.setWindowTitle(title)
        msg.setText(displayText)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def auto_login(self):
        conn = create_connection('ticket.db')
        if conn is not None:
            create_login_table(conn)
            cur = conn.cursor()
            cur.execute("SELECT * FROM login")
            row = cur.fetchone()

            if row is not None:
                if not self.isSessionExpired:
                    self.ui.username_field.setText(row[0])
                    self.ui.password_field.setText(row[1])
                else:
                    self.username = row[0]
                    self.password = row[1]
                self.login()
            conn.close()

    def startAnimation(self, animation):
        animation.start()

    def stopAnimation(self, animation):
        animation.stop()

    def enableAutoSearch(self):
        if not self.isAutoSearch:
            self.isAutoSearch = True
            self.ui.pushButton_7.setText("Disable Auto Search")
            self.ui.pushButton_7.setStyleSheet("background-color: rgb(0, 98, 163);\n"
                                               "color: rgb(0, 0, 0);")
        else:
            self.isAutoSearch = False
            self.ui.pushButton_7.setText("Enable Auto Search")
            self.ui.pushButton_7.setStyleSheet("background-color: rgb(0, 98, 163);\n"
                                               "color: rgb(240, 240, 240);")

    def login(self):
        username = ""
        password = ""
        if not self.isSessionExpired:
            username = self.ui.username_field.text()
            password = self.ui.password_field.text()
            self.login_gif = QMovie(get_path("gifs/294-1.gif"))
            self.ui.loading.setMovie(self.login_gif)
            self.startAnimation(self.login_gif)
            self.setbtngrey(self.ui.btn_login)
        else:
            username = self.username
            password = self.password
        self.loginthread = QThread()
        self.loginworker = LoginWorker(username, password)
        self.loginworker.moveToThread(self.loginthread)
        self.loginthread.started.connect(self.loginworker.run)
        self.loginworker.finished.connect(self.loginthread.quit)
        self.loginworker.finished.connect(self.loginworker.deleteLater)
        self.loginworker.result.connect(self.loginProgress)
        self.loginworker.error.connect(self.loginError)
        self.loginworker.session.connect(self.ticketSession)
        self.loginthread.finished.connect(self.loginthread.deleteLater)
        self.loginthread.start()

    @pyqtSlot(object)
    def ticketSession(self, _sess):
        self.session = _sess

    @pyqtSlot(int)
    def loginProgress(self, status):
        if status == 200:
            if not self.isSessionExpired:
                self.stopAnimation(self.login_gif)
                self.ui.stackedWidget.setCurrentWidget(self.ui.parser_page)
                self.disableTabs(False)
                self.ParserPage()
            else:
                self.startSearch()
        elif status == 403:
            self.stopAnimation(self.login_gif)
            self.setbtndefault(self.ui.btn_login)
            self.ui.loading.setWordWrap(True)
            self.displaymsg(
                self.ui.loading, 'Ticketing Forbidden. Please connect to Office Network')
            self.disableTabs()
            self.showDialog('Ticketing Error', QMessageBox.Critical,
                            'Ticketing Forbidden. Please connect to Office Network')
        elif status == 422:
            self.stopAnimation(self.login_gif)
            self.setbtndefault(self.ui.btn_login)
            self.ui.loading.setWordWrap(True)
            self.displaymsg(self.ui.loading,
                            'Username or Password is incorrect')
            self.disableTabs()
            self.showDialog("Login Error", QMessageBox.Critical,
                            'Username or Password is incorrect')

    @pyqtSlot(tuple)
    def loginError(self, error):
        self.ui.loading.setWordWrap(True)
        self.setbtndefault(self.ui.btn_login)
        self.displaymsg(self.ui.loading,
                        f'Unable to Connect. Please try again. {error[0]}')
        self.disableTabs()
        self.showDialog("Login Error", QMessageBox.Critical,
                        f'Unable to Connect. Please try again. {error[0]}')

    def ParserPage(self):
        self.ui.tableWidget.setHidden(True)
        self.ui.groupBox.setHidden(True)
        self.ui.groupBox_2.setHidden(True)
        self.ui.groupBox_9.setHidden(True)
        self.ui.tabWidget.setHidden(True)
        self.ui.empty_frame.setMinimumHeight(500)
        self.ui.empty_frame.setMaximumHeight(500)

    def PrepareParseWorker(self, module='mbox'):
        self.setbtngrey(self.ui.btn_parse_option)
        self.ui.toolButton.setEnabled(False)
        self.parsethread = QThread()
        self.parseworker = ParserWorker(self.EmlMBoxPath, module)
        self.parseworker.moveToThread(self.parsethread)
        self.parsethread.started.connect(self.parseworker.run)
        self.parseworker.finished.connect(self.parsethread.quit)
        self.parseworker.finished.connect(self.parseworker.deleteLater)
        self.parseworker.loading.connect(self.LoadingFiles)
        self.parseworker.error.connect(self.parseError)
        self.parseworker.progress.connect(self.ParserProgress)
        self.parseworker.progress_bar.connect(self.UpdateProgressBar)
        self.parseworker.result.connect(self.ParserResult)
        self.parseworker.finished.connect(
            lambda: self.ui.pushButton.setEnabled(True))
        self.parseworker.finished.connect(
            lambda: self.ui.btn_parse_clear.setHidden(False))
        self.parsethread.finished.connect(self.parsethread.deleteLater)
        self.parsethread.start()

    def parse(self):
        if self.ui.radio_eml.isChecked():
            if not self.ui.tableWidget.isHidden():
                self.ui.tableWidget.setRowCount(0)
                self.ConfigureTable(self.ui.tableWidget, self.cols)
            self.PrepareParseWorker('eml')
        elif self.ui.eml_mbox.isChecked():
            if not self.ui.tableWidget.isHidden():
                self.ui.tableWidget.setRowCount(0)
                self.ConfigureTable(self.ui.tableWidget, self.cols)
            self.PrepareParseWorker()

    def setbtngrey(self, btn, style="background-color: rgb(128, 128, 128);color: rgb(255, 255, 255);"):
        btn.setEnabled(False)
        btn.setStyleSheet(style)

    def setbtndefault(self, btn, style="background-color: rgb(0, 98, 163);color: rgb(250, 250, 250);"):
        btn.setEnabled(True)
        btn.setStyleSheet(style)

    def clearTable(self, table):
        table.setRowCount(0)
        self.emails = []
        self.ui.groupBox_9.setHidden(True)
        self.ui.groupBox_2.setHidden(True)
        self.ui.groupBox.setHidden(True)
        self.ui.tabWidget.setHidden(True)
        self.ui.tableWidget_2.setRowCount(0)
        self.ui.tableWidget_3.setRowCount(0)
        self.ui.tableWidget_7.setRowCount(0)
        self.ui.btn_parse_option.setEnabled(True)
        self.ui.toolButton.setEnabled(True)
        self.ui.btn_parse_clear.setHidden(True)
        self.setbtndefault(self.ui.btn_parse_option)
        conn = create_connection('ticket.db')
        if conn is not None:
            cur = conn.cursor()
            cur.execute('SELECT * FROM mbox')
            rows = cur.fetchall()
            if rows is not None:
                for row in rows:
                    if row[5] is None:
                        val = (row[0],)
                        cur.execute("DELETE FROM mbox WHERE rowid=?", val)
                cur.connection.commit()

    def ConfigureTable(self, table, cols, rowCount=None, colCount=4, matchCol=None):
        table.setColumnCount(colCount)
        if rowCount is None:
            # necessary even when there are no rows in the table
            table.setRowCount(1)
        else:
            table.setRowCount(rowCount)
        table.setSortingEnabled(False)
        table.setHorizontalHeaderLabels(cols)
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        table.verticalHeader().hide()
        table.setColumnWidth(0, 50)
        table.setColumnWidth(1, 150)
        if matchCol:
            table.setColumnWidth(2, 70)
            table.setColumnWidth(3, 300)
        else:
            table.setColumnWidth(2, 300)
        table.horizontalHeader().setStretchLastSection(True)

    def searchUnderway(self):
        if self.isSearch:
            self.ui.frame_6.setHidden(False)

    def updateRows(self, **kwargs):

        QTable = kwargs['table']
        isAttachment = kwargs['attachment']
        mode = kwargs['mode']
        colItems = kwargs['colItems']

        for i, colItem in enumerate(colItems):
            cellWidget = QTableWidgetItem()
            if mode == "allmails" and isAttachment and i == 2:
                cellWidget.setIcon(self.green_clip_icon)
            elif mode == "email" and isAttachment and i == 2:
                cellWidget.setIcon(self.green_clip_icon)
            elif mode == "subject" and isAttachment and i == 3:
                cellWidget.setIcon(self.green_clip_icon)
            elif mode == "find":
                if i == 2 and isAttachment:
                    cellWidget.setIcon(self.green_clip_icon)
                if i == 1:
                    linkTemplate = '<a href="{0}">{1}</a>'
                    self.label = HyperlinkLabel(QTable)
                    self.label.setText(
                        linkTemplate.format(colItem[1], colItem[0]))
            if mode == "find" and i == 1:
                QTable.setCellWidget(int(colItems[0])-1, i, self.label)
            else:
                if i == 1 and isinstance(colItem, list) and len(colItem) == 1:
                    cellWidget.setText(f"{colItem[0]}")
                else:
                    cellWidget.setText(f"{colItem}")
                QTable.setItem(int(colItems[0])-1, i, cellWidget)

    def updateTable(self, table, item, allMails=None, find=None, searching=None, db=None, subject=None):
        table.insertRow(int(item[0]))
        if searching and subject:
            item[2] = "%.2f" % float(item[2]) + "%"

        if allMails:
            attachment = True if len(item[4]) > 0 else False
            colItems = [col for col in item[:-1]]
            self.updateRows(mode="allmails", attachment=attachment,
                            colItems=colItems, table=table)
        elif find:
            attachment = True if item[5] > 0 else False
            colItems = [col for col in item[:-1]]
            self.updateRows(mode="find", attachment=attachment,
                            colItems=colItems, table=table)

        elif searching:
            colItems = [col for col in item[:-2]]
            if subject:
                attachment = True if len(item[5]) > 0 else False
                self.updateRows(mode="subject", attachment=attachment,
                                colItems=colItems, table=table)
            else:
                attachment = True if len(item[4]) > 0 else False
                self.updateRows(mode="email", attachment=attachment,
                                colItems=colItems, table=table)
        elif db:
            attach_string = item[4].strip("[]")

            if len(attach_string) > 0:
                attachment = True
            else:
                attachment = False

            if subject:
                colItems = [col for col in item[:-2]]
                colItems.insert(2, str("%.2f" % item[5])+"%")

                self.updateRows(mode="subject", attachment=attachment,
                                colItems=colItems, table=table)
            else:
                colItems = [col for col in item[:-1]]
                self.updateRows(mode="email", attachment=attachment,
                                colItems=colItems, table=table)

    @pyqtSlot(tuple)
    def LoadingFiles(self, data):
        self.ui.groupBox_9.setHidden(False)
        self.gif = QMovie(get_path("gifs/Dual Ring-1s-31px.gif"))
        if data[0] == 1:
            self.ui.loading_msg.setHidden(False)
            self.ui.loading_gif.setHidden(False)
            self.ui.loading_msg.setText(data[1])
            self.ui.loading_msg.setStyleSheet("color: rgb(0, 0, 0);")
            self.ui.loading_msg.setFont(self.font)
            self.ui.loading_gif.setMovie(self.gif)
            self.startAnimation(self.gif)
        elif data[0] == 2:
            self.stopAnimation(self.gif)
            self.total_emails = data[2]
            self.stopAnimation(self.gif)
            self.ui.loading_gif.setHidden(True)
            self.ui.loading_msg.setStyleSheet("color: rgb(0, 0, 0);")
            self.ui.loading_msg.setFont(self.font)
            self.ui.loading_msg.setText(data[1])
            if self.ui.tableWidget.isHidden():
                self.ui.tableWidget.setHidden(False)
                self.ConfigureTable(self.ui.tableWidget, self.cols)
            self.ui.empty_frame.setHidden(True)
            self.ui.groupBox.setHidden(False)
            self.ui.label.setHidden(False)
            self.ui.pushButton.setEnabled(False)

    @pyqtSlot(tuple)
    def parseError(self, error):
        self.ui.groupBox_9.setHidden(False)
        self.ui.loading_msg.setHidden(False)
        self.ui.loading_gif.setHidden(True)
        self.ui.loading_msg.setStyleSheet("color: rgb(255, 8, 12);")
        self.ui.loading_msg.setFont(self.font)
        self.ui.loading_msg.setText(error[0])
        self.ui.btn_parse_option.setEnabled(True)
        self.ui.btn_parse_option.setStyleSheet(
            "background-color: rgb(0, 98, 163);color: rgb(250, 250, 250);")
        self.showDialog("Warning", QMessageBox.Warning,
                        f"{error[0]} : {error[1]}")

    @pyqtSlot(list)
    def ParserProgress(self, email_item):
        self.updateTable(self.ui.tableWidget, email_item, allMails=True)

    @pyqtSlot(tuple)
    def UpdateProgressBar(self, incr):
        self.ui.progressBar.setValue(incr[0])
        self.ui.label.setText(
            f"Processed : [ {incr[1]} of {self.total_emails} ]")

    @pyqtSlot(list)
    def ParserResult(self, emails):
        self.emails = emails

    def Counter(self):
        self.time = self.time + 1

    def stopTimer(self):
        self.isSearch = False
        self.ui.frame_6.setHidden(True)
        self.timer.stop()

    def convertSecsToHHMMSS(self, seconds):
        seconds = seconds % (24 * 3600)
        hour = seconds // 3600
        seconds %= 3600
        minutes = seconds // 60
        seconds %= 60
        return "%d:%02d:%02d" % (hour, minutes, seconds)

    def startSearch(self):
        self.isSearch = True
        if not self.isSessionExpired:
            self.time = 0
            self.timer = QTimer()
            self.timer.start(1000)
            self.timer.timeout.connect(self.Counter)
            self.text_gif = QMovie(get_path('gifs/text_fading.gif'))
            self.ui.label_9.setMovie(self.text_gif)
            self.startAnimation(self.text_gif)
            self.ui.groupBox_2.setHidden(False)
            self.ui.groupBox.setHidden(True)
            self.ui.tabWidget.setHidden(False)
            self.ui.groupBox_9.setHidden(True)
            self.ui.btn_parse_clear.setEnabled(False)
            self.ConfigureTable(self.ui.tableWidget_2, self.cols)
            self.ConfigureTable(self.ui.tableWidget_3,
                                self.cols_sub_table, colCount=5, matchCol=True)
            self.ConfigureTable(self.ui.tableWidget_7, self.cols)

        self.searchthread = QThread()
        self.searchworker = SearchWorker(
            self.emails, self.session, self.startFrom)
        if self.isSessionExpired:
            self.isSessionExpired = False
            self.expiredAtCount = 0

        self.searchworker.moveToThread(self.searchthread)
        self.searchthread.started.connect(self.searchworker.run)
        self.searchworker.pBar.connect(self.UpdateSearchPBar)
        self.searchworker.processing.connect(self.SearchProcessing)
        self.searchworker.email_not_found.connect(
            self.updateEmailNotFoundTable)
        self.searchworker.subject_not_found.connect(
            self.updateSubjectNotFoundTable)
        self.searchworker.error.connect(self.searchError)
        self.searchworker.sessionError.connect(self.sessionExpired)
        self.searchworker.expiredAt.connect(self.setStartFrom)
        self.searchworker.warning.connect(self.Warning)
        self.searchworker.finished.connect(self.searchthread.quit)
        self.searchworker.finished.connect(self.searchworker.deleteLater)
        # self.searchworker.finished.connect(self.stopTimer)
        # self.searchworker.finished.connect(
        #     lambda: self.ui.btn_parse_clear.setEnabled(True))
        # self.searchworker.finished.connect(
        #     lambda: self.stopAnimation(self.text_gif))
        # self.searchworker.finished.connect(lambda: self.displaymsg(
        #     self.ui.label_9, text="Search Complete!", style="color: rgb(0, 98, 163);"))
        # self.searchworker.finished.connect(lambda: self.showDialog(
        #     "Search Job", QMessageBox.Information, "Search Complete!"))
        self.searchworker.finished.connect(self.searchComplete)
        self.searchthread.finished.connect(self.searchthread.deleteLater)
        self.searchthread.start()

    @pyqtSlot()
    def searchComplete(self):
        if not self.isSessionExpired:
            self.displaymsg(
                self.ui.label_9, text="Search Complete!", style="color: rgb(0, 98, 163);")
            self.showDialog(
                "Search Job", QMessageBox.Information, "Search Complete!")
            self.stopAnimation(self.text_gif)
            self.ui.btn_parse_clear.setEnabled(True)
            self.stopTimer()
        else:
            self.auto_login()

    @pyqtSlot(object)
    def setStartFrom(self, expiredAt):
        self.expiredAtCount += 1
        if self.expiredAtCount > 1:
            _dict = {"idx": 0, "email_idx": 0,
                     "subject_idx": 0, "err_idx": 0}
            for key, value in self.failedAt.items():
                if (key == "idx" and self.failedAt[key] > 0 and expiredAt[key] <= self.failedAt[key]):
                    if value > 0:
                        _dict[key] = expiredAt[key] - 1
                    else:
                        _dict[key] = expiredAt[key]
                elif (key == "idx" and self.failedAt[key] > 0 and expiredAt[key] >= self.failedAt[key]):
                    _dict[key] = self.failedAt[key] - 1
                elif (key == "idx" and self.failedAt[key] == 0):
                    if value > 0:
                        _dict[key] = expiredAt[key] - 1
                    else:
                        _dict[key] = expiredAt[key]
                elif (key == "email_idx" and self.failedAt[key] > 0 and expiredAt[key] <= self.failedAt[key]):
                    _dict[key] = self.failedAt[key]
                elif (key == "email_idx" and self.failedAt[key] > 0 and expiredAt[key] >= self.failedAt[key]):
                    _dict[key] = expiredAt[key]
                elif (key == "email_idx" and self.failedAt[key] == 0):
                    _dict[key] = expiredAt[key]
                elif (key == "subject_idx" and self.failedAt[key] > 0 and expiredAt[key] <= self.failedAt[key]):
                    _dict[key] = self.failedAt[key]
                elif (key == "subject_idx" and self.failedAt[key] > 0 and expiredAt[key] >= self.failedAt[key]):
                    _dict[key] = expiredAt[key]
                elif (key == "subject_idx" and self.failedAt[key] == 0):
                    _dict[key] = expiredAt[key]
                elif (key == "err_idx" and self.failedAt[key] > 0 and expiredAt[key] <= self.failedAt[key]):
                    _dict[key] = self.failedAt[key]
                elif (key == "err_idx" and self.failedAt[key] > 0 and expiredAt[key] >= self.failedAt[key]):
                    _dict[key] = expiredAt[key]
                elif (key == "err_idx" and self.failedAt[key] == 0):
                    _dict[key] = expiredAt[key]

            self.startFrom = deepcopy(_dict)
            self.failedAt = deepcopy(_dict)
        else:
            if expiredAt["idx"] > 0:
                expiredAt["idx"] = expiredAt["idx"] - 1
            self.startFrom = deepcopy(expiredAt)
            self.failedAt = deepcopy(expiredAt)

    @pyqtSlot(bool)
    def sessionExpired(self, expired):
        if expired:
            self.isSessionExpired = True

    @pyqtSlot(bool)
    def Warning(self, abnormal):
        if abnormal:
            self.showDialog('Search Error', QMessageBox.Warning,
                            "Search Terminated abnormally")

    @pyqtSlot(tuple)
    def SearchProcessing(self, ticket):
        self.ui.label_5.setText(f"Processing     : {ticket[0]} ")
        self.ui.label_6.setText(
            f"Processed      : [ {ticket[1]} of {len(self.emails)} ]")
        self.ui.label_7.setText(
            f"Time Elapsed   : {self.convertSecsToHHMMSS(self.time)}")

    @pyqtSlot(list)
    def updateEmailNotFoundTable(self, ticket):
        self.updateTable(self.ui.tableWidget_2, ticket, searching=True)

    @pyqtSlot(list)
    def updateSubjectNotFoundTable(self, ticket):
        self.updateTable(self.ui.tableWidget_3, ticket,
                         searching=True, subject=True)

    @pyqtSlot(int)
    def UpdateSearchPBar(self, n):
        self.ui.label_11.setText(f"Search Underway {n}%")
        self.ui.progressBar_2.setValue(n)

    @pyqtSlot(tuple)
    def searchError(self, error):
        self.updateTable(self.ui.tableWidget_7, error[1], searching=True)
        self.showDialog('Search Error', QMessageBox.Critical,
                        f'''{error[0]} : {error[1][1:]} : {error[2]}''')

    @pyqtSlot(list)
    def loadEmails(self, emails):
        temp = list(emails[1])  # convert tuple to list
        temp[0] = emails[0]
        self.updateTable(self.ui.tableWidget_4, temp, db=True)

    @pyqtSlot(list)
    def loadSubjects(self, subject):
        temp = list(subject[1])
        temp[0] = subject[0]
        self.updateTable(self.ui.tableWidget_5, temp, subject=True, db=True)

    def comboxSelected(self, index):
        if index >= 0:
            self.comboIndex = self.comboitems[index][0]
        else:
            if len(self.comboitems) > 0:
                self.comboIndex = self.comboitems[0][0]
        self.ui.pushButton_3.setEnabled(True)

    def loadPreviousResults(self):
        self.ui.tableWidget_4.setRowCount(0)  # clear table
        self.ConfigureTable(self.ui.tableWidget_4, self.cols)
        self.ui.tableWidget_5.setRowCount(0)  # clear table
        self.ConfigureTable(self.ui.tableWidget_5,
                            self.cols_sub_table, colCount=5, matchCol=True)

        if self.comboIndex is not None:
            self.loadthread = QThread()
            self.loadworker = LoadWorker(self.comboIndex)
            self.loadworker.moveToThread(self.loadthread)
            self.loadworker.finished.connect(self.loadthread.quit)
            self.loadworker.finished.connect(self.loadworker.deleteLater)
            self.loadworker.email_not_found.connect(self.loadEmails)
            self.loadworker.subject_not_found.connect(self.loadSubjects)
            self.loadthread.started.connect(self.loadworker.run)
            self.loadthread.finished.connect(self.loadthread.deleteLater)
            self.loadthread.start()

    def loadComboBox(self):
        self.comboitems = []
        self.ui.comboBox.clear()
        conn = create_connection('ticket.db')
        with conn:
            if conn is not None:
                with conn:
                    query = '''SELECT * FROM mbox'''
                    cur = conn.cursor()
                    cur.execute(query)
                    rows = cur.fetchall()
                    if rows is not None:
                        for row in rows:
                            if row[5] is not None:
                                self.comboitems.append([row[0], row[1]])
                                self.ui.comboBox.addItem(f"{row[0]}. {row[1]}")

    def autoSearch(self, table):
        if self.isAutoSearch:
            for item in table.selectedItems():
                if item.column() == 1 and not self.isSearching:
                    query = item.text()
                    self.ui.lineEdit.setText(query)
                    self.querySearch(query)

    def querySearch(self, query=None):
        self.isSearching = True
        if not query:
            query = self.ui.lineEdit.text()
        self.ui.tableWidget_6.setRowCount(0)  # clear table
        self.ui.pushButton_2.setEnabled(False)
        self.ui.gif_search.setHidden(False)
        self.searchGif = QMovie(get_path("gifs/Wedges-2.1s-25px.gif"))
        self.ui.gif_search.setMovie(self.searchGif)
        self.startAnimation(self.searchGif)
        column = ['SI NO', "Ticket", "Subject", "Status", "Last Updated"]
        self.ConfigureTable(self.ui.tableWidget_6, column, colCount=5)
        self.queryThread = QThread()
        self.queryWorker = QuerySearchWorker(self.session, query)
        self.queryWorker.moveToThread(self.queryThread)
        self.queryWorker.finished.connect(self.queryThread.quit)
        self.queryWorker.finished.connect(self.queryWorker.deleteLater)
        self.queryWorker.result.connect(self.searchResult)
        self.queryWorker.finished.connect(
            lambda: self.stopAnimation(self.searchGif))
        self.queryWorker.finished.connect(
            lambda: self.ui.gif_search.setHidden(True))
        self.queryWorker.finished.connect(
            lambda: self.ui.pushButton_2.setEnabled(True))
        self.queryThread.started.connect(self.queryWorker.run)
        self.queryThread.finished.connect(self.queryThread.deleteLater)
        self.queryThread.start()

    @pyqtSlot(object)
    def searchResult(self, result):
        if result is not None:
            self.updateTable(self.ui.tableWidget_6, result, find=True)
        self.isSearching = False


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # clipboard = app.clipboard()
    window = MainApp()
    window.show()
    sys.exit(app.exec())
