import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from PyQt5 import uic, QtCore
import CompareFiles
import os

Ui_MainWindow, QtMainBaseClass = uic.loadUiType('DatabaseComparerGui.ui')
Ui_ListWindow, QtListBaseClass = uic.loadUiType('BlackListGui.ui')
ignore_list_path_init = 'ignore_vilkar.txt'
ignore_list_path_cpr = 'ignore_hk.txt'


class EmittingStream(QtCore.QObject):
    textWritten = QtCore.pyqtSignal(str)

    def write(self, text):
        self.textWritten.emit(str(text))

    def flush(self):
        pass


class ListWindow(QMainWindow):
    def __init__(self, ignore_file, parent):
        super().__init__()
        self.ui = Ui_ListWindow()
        self.ui.setupUi(self)
        self.current_ignore_file = ignore_file
        self.parent = parent

        # Assign functions to ui elements
        self.ui.saveButton.clicked.connect(self.save_selection)

    def save_selection(self):
        selected_items = [i.text().split(', ')[0] for i in self.ui.mainList.selectedItems()]
        with open(self.current_ignore_file, 'a') as f:
            f.write('\n'.join(selected_items) + '\n')
        self.close()
        with open(self.current_ignore_file, 'r') as f:
            ignore_list = [l.rstrip() for l in f.readlines()]
        if self.current_ignore_file == ignore_list_path_init:
            self.parent.vilkar_ignore = ignore_list
        elif self.current_ignore_file == ignore_list_path_cpr:
            self.parent.hk_ignore = ignore_list


class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Fields
        self.db_file_path = ''
        self.vilkar_file_path = ''
        self.hk_file_path = ''
        self.emails_file_path = ''
        self.last_dir = ''
        self.database = None
        self.medlem = None
        self.vilkar = None
        self.hk = None
        self.maillist = None
        self.list_dialog = None
        self.vilkar_ignore = []
        self.hk_ignore = []
        if os.path.isfile(ignore_list_path_init):
            with open(ignore_list_path_init, 'r') as f:
                self.vilkar_ignore = [l.rstrip() for l in f.readlines()]
        if os.path.isfile(ignore_list_path_cpr):
            with open(ignore_list_path_cpr, 'r') as f:
                self.hk_ignore = [l.rstrip() for l in f.readlines()]

        # Install the custom output stream
        sys.stdout = EmittingStream(textWritten=self.normal_output_written)

        # Assign functions to ui elements
        self.ui.clearButton.clicked.connect(self.clear_output)
        self.ui.loadDb.clicked.connect(self.select_db)
        self.ui.loadVilkar.clicked.connect(self.select_vilkar)
        self.ui.loadHk.clicked.connect(self.select_hk)
        self.ui.loadEmails.clicked.connect(self.select_emails)
        self.ui.compareInit.clicked.connect(lambda: self.compare_init(print_out=True))
        self.ui.compareCpr.clicked.connect(lambda: self.compare_cpr(print_out=True))
        self.ui.compareEmails.clicked.connect(lambda: self.compare_emails(print_out=True))
        self.ui.blacklistInit.clicked.connect(lambda: self.blacklist(ignore_file=ignore_list_path_init))
        self.ui.blacklistCpr.clicked.connect(lambda: self.blacklist(ignore_file=ignore_list_path_cpr))

    def blacklist(self, ignore_file):
        if ignore_file == ignore_list_path_init:
            db_only, file_only = self.compare_init(print_out=False)
            display_info = [e + ', ' + self.vilkar['Job'][self.vilkar['Initials'].str.upper() == e].iloc[0] for e in
                            file_only]
        elif ignore_file == ignore_list_path_cpr:
            db_only, file_only = self.compare_cpr(print_out=False)
            display_info = [e + ', ' + self.hk['Medlemskategori'][self.hk['CPR'].str.replace('-', '') == e].iloc[0] for
                            e in file_only]
        self.list_dialog = ListWindow(ignore_file=ignore_file, parent=self)
        self.list_dialog.ui.mainList.addItems(display_info)
        self.list_dialog.show()

    def __del__(self):
        # Restore sys.stdout
        sys.stdout = sys.__stdout__

    def clear_output(self):
        self.ui.textConsole.setText('')

    def select_db(self):
        self.db_file_path, _ = QFileDialog.getOpenFileName(self, "Load database file", self.last_dir,
                                                           "Excel Files (*.xlsx)")
        self.last_dir = os.path.dirname(self.db_file_path)
        self.ui.textDb.setText(self.db_file_path)
        print('Database set to: ' + self.db_file_path)

    def select_vilkar(self):
        self.vilkar_file_path, _ = QFileDialog.getOpenFileName(self, "Load vilkar file", self.last_dir,
                                                               "Excel Files (*.xlsx)")
        self.last_dir = os.path.dirname(self.vilkar_file_path)
        self.ui.textVilkar.setText(self.vilkar_file_path)
        print('Vilkar set to: ' + self.vilkar_file_path)

    def select_hk(self):
        self.hk_file_path, _ = QFileDialog.getOpenFileName(self, "Load hk file", self.last_dir, "Excel Files (*.xlsx)")
        self.last_dir = os.path.dirname(self.hk_file_path)
        self.ui.textHk.setText(self.hk_file_path)
        print('HK set to: ' + self.hk_file_path)

    def select_emails(self):
        self.emails_file_path, _ = QFileDialog.getOpenFileName(self, "Load email file", self.last_dir,
                                                               "Text Files (*.txt)")
        self.last_dir = os.path.dirname(self.emails_file_path)
        self.ui.textEmails.setText(self.emails_file_path)
        print('Emails set to: ' + self.emails_file_path)

    def compare_init(self, print_out):
        self.database, self.medlem = CompareFiles.read_db(self.db_file_path)
        self.vilkar = CompareFiles.read_vilkar(self.vilkar_file_path)
        return CompareFiles.compare_vilkar(self.database, self.vilkar, self.vilkar_ignore, print_out)

    def compare_cpr(self, print_out):
        self.database, self.medlem = CompareFiles.read_db(self.db_file_path)
        self.hk = CompareFiles.read_hk(self.hk_file_path)
        return CompareFiles.compare_hk(self.database, self.hk, self.hk_ignore, print_out)

    def compare_emails(self, print_out):
        self.database, self.medlem = CompareFiles.read_db(self.db_file_path)
        self.maillist = CompareFiles.read_maillist(self.emails_file_path)
        return CompareFiles.compare_maillist(self.database, self.medlem, self.maillist, print_out)

    def normal_output_written(self, text):
        """ Append text to the QTextEdit """
        self.ui.textConsole.append(text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())

# UI designed with qt designer
# conda prompt: designer
#
# Exe build instructions:
# create -n py35 python=3.5
# activate py35
# conda install -c acellera -n py35 pyinstaller
# pyinstaller --onefile --windowed --icon=database_refresh.ico yourprogram.py
