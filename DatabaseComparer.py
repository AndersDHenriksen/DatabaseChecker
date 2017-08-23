import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from PyQt5 import uic, QtCore
import CompareFiles
import os

qtCreatorFile = 'DatabaseComparerGui.ui'  # Enter file here.
Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)


class EmittingStream(QtCore.QObject):
    textWritten = QtCore.pyqtSignal(str)

    def write(self, text):
        self.textWritten.emit(str(text))

    def flush(self):
        pass


class MyApp(QMainWindow):
    def __init__(self):
        super(MyApp, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle('DatebasesComparer')

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

        # Install the custom output stream
        sys.stdout = EmittingStream(textWritten=self.normalOutputWritten)

        # Assign functions to ui elements
        self.ui.clearButton.clicked.connect(self.clear_output)
        self.ui.loadDb.clicked.connect(self.select_db)
        self.ui.loadVilkar.clicked.connect(self.select_vilkar)
        self.ui.loadHk.clicked.connect(self.select_hk)
        self.ui.loadEmails.clicked.connect(self.select_emails)
        self.ui.compareInit.clicked.connect(self.compare_init)
        self.ui.compareCpr.clicked.connect(self.compare_cpr)
        self.ui.compareEmails.clicked.connect(self.compare_emails)

    def __del__(self):
        # Restore sys.stdout
        sys.stdout = sys.__stdout__

    def clear_output(self):
        self.ui.textConsole.setText('')

    def select_db(self):
        self.db_file_path, _ = QFileDialog.getOpenFileName(self, "Load database file", self.last_dir, "Excel Files (*.xlsx)")
        self.last_dir = os.path.dirname(self.db_file_path)
        self.ui.textDb.setText(self.db_file_path)
        print('Database set to: ' + self.db_file_path)

    def select_vilkar(self):
        self.vilkar_file_path, _ = QFileDialog.getOpenFileName(self, "Load vilkar file", self.last_dir, "Excel Files (*.xlsx)")
        self.last_dir = os.path.dirname(self.vilkar_file_path)
        self.ui.textVilkar.setText(self.vilkar_file_path)
        print('Vilkar set to: ' + self.vilkar_file_path)

    def select_hk(self):
        self.hk_file_path, _ = QFileDialog.getOpenFileName(self, "Load hk file", self.last_dir, "Excel Files (*.xlsx)")
        self.last_dir = os.path.dirname(self.hk_file_path)
        self.ui.textHk.setText(self.hk_file_path)
        print('HK set to: ' + self.hk_file_path)

    def select_emails(self):
        self.emails_file_path, _ = QFileDialog.getOpenFileName(self, "Load email file", self.last_dir, "Text Files (*.txt)")
        self.last_dir = os.path.dirname(self.emails_file_path)
        self.ui.textEmails.setText(self.emails_file_path)
        print('Emails set to: ' + self.emails_file_path)

    def compare_init(self):
        self.database, self.medlem = CompareFiles.read_db(self.db_file_path)
        self.vilkar = CompareFiles.read_vilkar(self.vilkar_file_path)
        CompareFiles.compare_vilkar(self.database, self.vilkar)

    def compare_cpr(self):
        self.database, self.medlem = CompareFiles.read_db(self.db_file_path)
        self.hk = CompareFiles.read_hk(self.hk_file_path)
        CompareFiles.compare_hk(self.database, self.hk)

    def compare_emails(self):
        self.database, self.medlem = CompareFiles.read_db(self.db_file_path)
        self.maillist = CompareFiles.read_maillist(self.emails_file_path)
        CompareFiles.compare_maillist(self.database, self.medlem, self.maillist)

    def normalOutputWritten(self, text):
        """ Append text to the QTextEdit """
        self.ui.textConsole.append(text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())