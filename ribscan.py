#!/usr/bin/python
"""
Scan PDF using NAPS2 console and attach to Thunderbird email application for Dave Schwab

NAPS2 is a .NET application only available for Windows. Would like to port to .NET Core for linux.
SimpleScan and XSane are so focused on UI that they do not support scanning from the command line.

Ugly 2 line batch file solution:
C:/Program Files (x86)/NAPS2/naps2.console.exe -o C:/temp/out.pdf
C:/Program Files (x86)/Mozilla Thunderbird/thunderbird.exe -compose attachment='C:/temp/out.pdf'
"""

import sys
import datetime
import os.path

from subprocess import run, PIPE
from configparser import ConfigParser

from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMessageBox

from ui_ribscan import Ui_MainWindow

class RibscanAppDefaults():
    """
    RibscanAppDefaults - The hard coded default settings used to initialize the application
    configparser/ini settings which can then be user-overridden on the Settings tab.  These
    defaults will be used when the ini file is missing.
    """

    ini_name = 'ribscan.ini'
    icon_name = 'icon.png'

    # The default PDF folder can be set to the user's home folder, the CWD, or a custom default
    path_pdf_folder = os.path.expanduser('~')
    path_pdf_folder_ribscan = os.path.abspath('.')
    path_pdf_folder_dave = os.path.normpath('C:\\Users\\Dave\\Desktop\\attachments')

    path_thunderbird_bin_windows = os.path.normpath('C:\\Program Files (x86)\\Mozilla Thunderbird\\thunderbird')
    path_thunderbird_bin_linux = os.path.normpath('/usr/bin/thunderbird')

    path_naps2_bin_windows = os.path.normpath('C:\\Program Files (x86)\\NAPS2\\naps2.console')
    path_naps2_bin_linux = os.path.abspath('./naps2.console.sh') # Fake NAPS2 for development

    def __init__(self):
        # Uncomment one of the 3 lines below (or use the GUI to change your folder preference)

        #self.path_pdf_folder = self.path_pdf_folder        # Current user's home folder
        #self.path_pdf_folder = self.path_pdf_folder_dave   # Custom user location
        self.path_pdf_folder = self.path_pdf_folder_ribscan # Ribscan folder (current working dir)

class WorkerTask(QtCore.QThread):
    """
    The scan and email operations take an indeterminate amount of time to complete and their
    progress cannot be measured. This means these operations will have to run in their own QThread.
    Signals must be defined as class variables.
    """
    taskFinished = QtCore.pyqtSignal(int, str, str)
    cmd_list = list()

    def set_cmd_list(self, items):
        self.cmd_list.clear()
        for item in items:
            self.cmd_list.append(item)

    def run(self):
        if sys.platform == "win32":
            cp = run([self.cmd_list[0], self.cmd_list[1], self.cmd_list[2]], stdout=PIPE, stderr=PIPE, shell=True)
        else:
            cp = run([self.cmd_list[0], self.cmd_list[1], self.cmd_list[2]], stdout=PIPE, stderr=PIPE)

        self.taskFinished.emit(cp.returncode, cp.stdout.decode('UTF-8'), cp.stderr.decode('UTF-8'))

class RibscanApp(QtWidgets.QMainWindow, Ui_MainWindow):
    """
    RibscanApp - Main application class
    """

    defaults = RibscanAppDefaults()
    config = ConfigParser()

    ini_folder = os.path.abspath('.')
    ini_file = os.path.join(ini_folder, defaults.ini_name)

    def init_settings(self):
        pdf_loc = self.defaults.path_pdf_folder

        if sys.platform == "linux" or sys.platform == "linux2":
            naps2_cmd = self.defaults.path_naps2_bin_linux
            thunderbird_cmd = self.defaults.path_thunderbird_bin_linux
        elif sys.platform == "win32":
            naps2_cmd = self.defaults.path_naps2_bin_windows
            thunderbird_cmd = self.defaults.path_thunderbird_bin_windows
        else:
            # Assume (probably incorrectly), that these commands are in the system path
            naps2_cmd = "naps2.console"
            thunderbird_cmd = "thunderbird"

        if os.path.exists(self.ini_file):
            self.config.read(self.ini_file)
        else:
            self.config['DEFAULT'] = {'path_pdf_folder': pdf_loc, \
                                      'path_naps2_command': naps2_cmd, \
                                      'path_thunderbird_command': thunderbird_cmd}

            with open(self.ini_file, 'w') as configfile:
                self.config.write(configfile)

    def set_pdf_filename(self):
        """
        Generate a unique filename in the format "Year-Month-Day_Hour.Minute.Second.pdf".
        This needs to be called each time the "Scan and Email PDF" button is clicked.
        """
        new_filename = '{0:%Y-%m-%d_%H.%M.%S}'.format(datetime.datetime.now()) + '.pdf'
        self.pdf_filename = os.path.join(self.config['DEFAULT']['path_pdf_folder'], new_filename)

    def toolButton_pdf_clicked(self):
        dialog = QtWidgets.QFileDialog()
        options = dialog.Options()
        options |= QtWidgets.QFileDialog.ShowDirsOnly
        dialog.setDirectory(self.config['DEFAULT']['path_pdf_folder'])
        dirname = dialog.getExistingDirectory(self, 'Choose a Folder')
        if dirname:
            goodpath = os.path.normpath(dirname)
            self.lineEdit_pdf.setText(goodpath)
            self.pushButton_save.setEnabled(True)

    def toolButton_naps2_clicked(self):
        dialog = QtWidgets.QFileDialog()
        options = dialog.Options()
        dirname, filename = os.path.split(self.config['DEFAULT']['path_naps2_command'])
        dialog.setDirectory(dirname)
        title = self.config['DEFAULT']['path_naps2_command']
        filename, _ = dialog.getOpenFileName(self, title, "","All Files (*)", options=options)
        if filename:
            goodpath = os.path.normpath(filename)
            self.lineEdit_naps2.setText(goodpath)
            self.pushButton_save.setEnabled(True)

    def toolButton_thunderbird_clicked(self):
        dialog = QtWidgets.QFileDialog()
        options = dialog.Options()
        dirname, filename = os.path.split(self.config['DEFAULT']['path_thunderbird_command'])
        dialog.setDirectory(dirname)
        title = self.config['DEFAULT']['path_thunderbird_command']
        filename, _ = dialog.getOpenFileName(self, title, "","All Files (*)", options=options)
        if filename:
            goodpath = os.path.normpath(filename)
            self.lineEdit_thunderbird.setText(goodpath)
            self.pushButton_save.setEnabled(True)

    def pushButton_saved_clicked(self):
        self.config['DEFAULT']['path_pdf_folder'] = self.lineEdit_pdf.text()
        self.config['DEFAULT']['path_naps2_command'] = self.lineEdit_naps2.text()
        self.config['DEFAULT']['path_thunderbird_command'] = self.lineEdit_thunderbird.text()

        with open(self.ini_file, 'w') as configfile:
            self.config.write(configfile)

        self.statusBar().showMessage("Configuration saved!")

    def scan2pdf_and_email(self):
        self.pushButton.setEnabled(False)
        self.set_pdf_filename()
        self.naps2_scan_to_pdf(self.pdf_filename)

    def naps2_scan_to_pdf(self, pdfname):
        cmd = (self.config['DEFAULT']['path_naps2_command'], '-o', pdfname)
        self.scantask.set_cmd_list(cmd)
        self.start_scan_task()

    def start_scan_task(self):
        self.progressBarScan.setRange(0, 0)
        self.statusBar().showMessage("Creating " + self.pdf_filename)
        self.scantask.start()
    
    def finished_scan_task(self, retint, outstr, errstr):
        self.progressBarScan.setRange(1, 1) # Stop progressbar pulse
        self.statusBar().clearMessage() # Clear the statusbar

        # NAPS2 logs errors to stdout only, 1 per line.  Return code is always zero.
        if errstr or retint:
            pass
        if outstr:
            for errormsg in outstr.splitlines():
                if errormsg:
                    QMessageBox.critical(self, 'Scanning Error', errormsg)
            self.finished_email_task()
        else:
            self.thunderbird_compose_with_attachment(self.pdf_filename) # Email attachment

    def thunderbird_compose_with_attachment(self, pdfname):
        cmd = (self.config['DEFAULT']['path_thunderbird_command'], '-compose', "attachment='" + pdfname + "'")
        self.emailtask.set_cmd_list(cmd)
        self.start_email_task()

    def start_email_task(self):
        self.progressBarEmail.setRange(0, 0)
        self.statusBar().showMessage("Switch to the Write: window to Send your Email attachment...")
        self.emailtask.start()

    def finished_email_task(self):
        self.progressBarEmail.setRange(1, 1) # Stop progressbar pulse
        self.pushButton.setEnabled(True) # Enable the main button
        self.statusBar().clearMessage() # Clear the statusbar

    def about_call(self):
        message = 'Scan document to PDF and attach to new Thunderbird E-Mail.\n\n'
        message += '(C) 2017 Andrew Logue.  Icon (C) 2013 Thomas Tamblyn.\n'
        message += 'Made for Dave Schwab at Ribstone Resources Ltd.'
        QMessageBox.about(self, 'About Ribscan', message)

    def settings_call(self):
        self.tabWidget.setCurrentIndex(1)

    def exit_call(self):
        sys.exit()

    def __init__(self):
        super(RibscanApp, self).__init__()

        self.scantask = WorkerTask()
        self.scantask.taskFinished.connect(self.finished_scan_task)

        self.emailtask = WorkerTask()
        self.emailtask.taskFinished.connect(self.finished_email_task)

        # Set the application icon
        self.setWindowIcon(QIcon(os.path.join(os.path.abspath('.'), self.defaults.icon_name)))

        # Set up the user interface from Qt Designer
        self.setupUi(self)

        # Initialize settings
        self.init_settings()

        # Main menu action hooks
        self.actionExit.triggered.connect(self.exit_call)
        self.actionSettings.triggered.connect(self.settings_call)
        self.actionAbout.triggered.connect(self.about_call)

        # Button action hooks
        self.toolButton_pdf.clicked.connect(self.toolButton_pdf_clicked)
        self.toolButton_naps2.clicked.connect(self.toolButton_naps2_clicked)
        self.toolButton_thunderbird.clicked.connect(self.toolButton_thunderbird_clicked)
        self.pushButton_save.clicked.connect(self.pushButton_saved_clicked)
        self.pushButton.clicked.connect(self.scan2pdf_and_email)

        # Populate UI items on Settings tab
        self.lineEdit_pdf.setText(self.config['DEFAULT']['path_pdf_folder'])
        self.lineEdit_naps2.setText(self.config['DEFAULT']['path_naps2_command'])
        self.lineEdit_thunderbird.setText(self.config['DEFAULT']['path_thunderbird_command'])

        # self.pdf_filename needs to be regenerated prior to each scan
        new_filename = '{0:%Y-%m-%d_%H.%M.%S}'.format(datetime.datetime.now()) + '.pdf'
        self.pdf_filename = os.path.join(self.config['DEFAULT']['path_pdf_folder'], new_filename)

        # Select the main tab
        self.tabWidget.setCurrentIndex(0)
        self.textBrowser.append("Load papers into the scanner and click the button to begin")

if __name__ == "__main__":
    APP = QtWidgets.QApplication(sys.argv)
    MAINWIN = RibscanApp()
    MAINWIN.show()
    sys.exit(APP.exec_())
