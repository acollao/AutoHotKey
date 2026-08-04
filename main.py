from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from sys import argv, exit
from wget import download
from shutil import rmtree, copy
from zipfile import ZipFile
from numpy import random

import pandas as pd
import ocr_tesseract_update

import PyPDF2
import ctypes, os, sys, json, time, win32gui, win32con, win32api, getpass
import atexit, signal, datetime, joblib, configparser, traceback
import win32com.client as win32

class PathwayWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.resize(800, 500)
        layout = QVBoxLayout()
        self.pathway_table = QTableWidget()
        self.pathway_table.setStyleSheet("QTableView { font-size: 14px; } QHeaderView::section { background-color: rgb(210, 210, 210); font-size: 13px; font-weight: bold; border: 0; border-right: 1px solid rgb(200, 200, 200); border-bottom: 1px solid rgb(200, 200, 200);}")
        layout.addWidget(self.pathway_table)
        self.setLayout(layout)
        self.setWindowTitle("Bill Rules")
        self.setWindowIcon(QIcon("icons\\icon.ico"))

class ImageViewerWindow(QMainWindow):
    def __init__(self, main_ui):
        super(ImageViewerWindow, self).__init__()
        uic.loadUi("ui\\image_viewer.ui", self)
        self.setWindowTitle("PDF Viewer")
        self.setWindowIcon(QIcon("icons\\icon.ico"))

        self.label = self.findChild(QLabel, "label")
        self.imgs_arr = os.listdir("resources\\image_results")
        self.img_index = 0
        self.main_ui = main_ui

        self.img_viewer = self.findChild(QLabel, "label")
        self.search_frame = self.findChild(QFrame, "frame")
        self.tbx_search = self.findChild(QLineEdit, "lineEdit")
        self.btn_search = self.findChild(QPushButton, "pushButton")
        self.prev_button = self.findChild(QPushButton, "pushButton_2")
        self.next_button = self.findChild(QPushButton, "pushButton_3")
        self.search_button = self.findChild(QPushButton, "pushButton_4")
        self.close_button = self.findChild(QPushButton, "pushButton_5")

        self.search_frame.setVisible(False)

        self.prev_button.clicked.connect(self.previous_image)
        self.next_button.clicked.connect(self.next_image)
        self.btn_search.clicked.connect(self.search_text)
        self.tbx_search.returnPressed.connect(self.search_text)
        self.search_button.clicked.connect(self.searchbutton_event)
        self.close_button.clicked.connect(self.searchbutton_event)

        QFontDatabase.addApplicationFont(f"{os.getcwd()}/resources/fonts/Aptos-Display.ttf")

    def searchbutton_event(self):
        if self.search_frame.isVisible():
            self.search_frame.setVisible(False)
            self.search_button.setVisible(True)
        else:
            self.search_frame.setVisible(True)
            self.search_button.setVisible(False)

    def search_text(self):
        self.search_frame.setVisible(False)
        self.search_button.setVisible(True)

        if self.tbx_search.text() == "":
            return 0

        ocr_tesseract_update.get_ocr_result(self.main_ui.op_report_pdf, self.tbx_search.text())
        self.load_image()

    def load_image(self):
        self.imgs_arr = os.listdir("resources\\image_results")
        
        # Load and display the current image
        if 0 <= self.img_index < len(self.imgs_arr):
            image_path = f"resources\\image_results\\{self.imgs_arr[self.img_index]}"
            pixmap = QPixmap(image_path)
            self.label.setPixmap(pixmap)
            self.label.setScaledContents(True)

        self.setWindowTitle(f"PDF Viewer - {self.imgs_arr[self.img_index]}")

        self.raise_()
        self.activateWindow()

    def previous_image(self):
        # Show the previous image
        if self.img_index > 0:
            self.img_index -= 1
            self.load_image()

    def next_image(self):
        # Show the next image
        page_length = len(self.imgs_arr)

        if self.img_index < page_length - 1:
            self.img_index += 1
            self.load_image()

class Info1Window(QMainWindow):
    def __init__(self, driver, main_ui):
        super(Info1Window, self).__init__()
        uic.loadUi("ui\\info1.ui", self)
        self.setWindowIcon(QIcon("icons\\icon.ico"))
        self.comboBox_selection = self.findChild(QComboBox, "comboBox_selection")
        self.lineEdit_actual = self.findChild(QLineEdit, "lineEdit_actual")
        self.pushButton_info1 = self.findChild(QPushButton, "pushButton_info1")
        self.driver = driver
        self.main_ui = main_ui

        self.comboBox_selection.currentIndexChanged.connect(self.select_info1)
        self.pushButton_info1.clicked.connect(self.submit_info1)

    def select_info1(self):
        self.lineEdit_actual.setText(self.comboBox_selection.currentText())

    def submit_info1(self):
        if not self.driver is None:
            txt = self.lineEdit_actual.text()
            if not txt.strip() == "":
                try:
                    UI.window_focus(self, "SmartCoder")
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame('menuFrame')
                    self.driver.execute_script("javascript:onMenuDemographic()")

                    time.sleep(3)
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame("applFrame")
                    try:
                        self.driver.find_element(By.ID, "optRow5").send_keys(txt.upper())
                        self.driver.find_element(By.ID, "optRow5").send_keys(Keys.ENTER)
                        time.sleep(2)
                    except:
                        UI.ui_to_front(self)
                        UI.statusbar_update(self.main_ui, "Something is wrong. Please enter Info1 manually.")
                except:
                    UI.ui_to_front(self)
                    UI.statusbar_update(self.main_ui, "Please enter claim number on Starataware.")
        else:
            UI.ui_to_front(self)
            UI.statusbar_update(self.main_ui, "Please connect to the browser.")

class UI(QMainWindow):
    def __init__(self):
        super(UI, self).__init__()
        uic.loadUi("ui\\main_v2.ui", self)
        self.setWindowIcon(QIcon("icons\\icon.ico"))

        self.config = configparser.ConfigParser()
        self.config.read("config.ini")

        # UI Initialization
        self.driver = None
        self.pathway_window = None
        self.imageviewer_window = None
        self.info1_window = None
        self.strataware_url = "https://strataware.lmig.com/"
        self.tmp_strata_url = ""
        self.username = win32api.GetUserName().lower()
        self.op_report_pdf = ""
        self.allocation = []
        self.pathway_rules = []
        self.services_array = []
        self.finished_allocations = []
        self.ocr_search_array = []
        self.cpt_array = []
        self.dup_array = []
        self.cpt_details = []
        self.cpt_conflicts = []
        self.checker = 0
        self.time = QTime(0, 0, 0)
        self.start_timer = False

        self.documents_folder = os.path.expanduser(f"~\Documents")
        self.fso = win32.Dispatch("Scripting.FileSystemObject")
        self.edge_version = self.fso.GetFileVersion("c:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")
        self.webd_version = 0

        

        # Randomize a port number to be used by Selenium web driver. 
        # Don't want to hard code a port number that will be used by multiple users.
        # This is a temporary workaround for the EncoderPro login issue.
        # Not sure if this is a permanent workaround, but it looks like it will be. Crossing fingers.
        self.port = random.randint(low=10000, high=30000)

        self.qtimer = QTimer()
        self.qtimer.setInterval(750)
        self.qtimer.timeout.connect(self.high_risk_notif)
        self.qtimer.start()

        self.ttimer = QTimer()
        self.ttimer.setInterval(1000)
        self.ttimer.timeout.connect(self.timer_event)
        self.ttimer.start()

        # Frames
        self.main_frame = self.findChild(QFrame, "frame_6")
        self.frm_diagidcx = self.findChild(QFrame, "frm_diagicdx_more")

        ## Error Propensity Frame
        self.lbl_icn = self.findChild(QLabel, "lbl_icn")
        self.lbl_claimnumber = self.findChild(QLabel, "lbl_claimnumber")
        self.lbl_holdgroup = self.findChild(QLabel, "lbl_holdgroup")
        self.lbl_jurisdiction = self.findChild(QLabel, "lbl_jurisdiction")
        self.lbl_riskstatus = self.findChild(QLabel, "lbl_riskstatus")
        self.lbl_billingtype = self.findChild(QLabel, "lbl_billingtype")
        self.lbl_totalcharges = self.findChild(QLabel, "lbl_totalcharges")
        self.lbl_totalallowance = self.findChild(QLabel, "lbl_totalallowance")
        self.lbl_errorpropensity = self.findChild(QLabel, "lbl_errorpropensity")
        # self.lbl_atriskdays = self.findChild(QLabel, "lbl_atriskdays")
        # self.lbl_criticaldays = self.findChild(QLabel, "lbl_criticaldays")
        # self.lbl_projatriskdate = self.findChild(QLabel, "lbl_projatriskdate")
        # self.lbl_projcritdate = self.findChild(QLabel, "lbl_projcritdate")
        # self.lbl_priority = self.findChild(QLabel, "lbl_priority")
        # self.lbl_inventorydate = self.findChild(QLabel, "lbl_inventorydate")
        # self.lbl_incomingleftover = self.findChild(QLabel, "lbl_incomingleftover")

        ## Demographics
        self.tbx_claimant = self.findChild(QLabel, "tbx_claimant")
        self.tbx_idcx1 = self.findChild(QLabel, "tbx_idcx1")
        self.tbx_idcx2 = self.findChild(QLabel, "tbx_idcx2")
        self.tbx_idcx3 = self.findChild(QLabel, "tbx_idcx3")
        self.tbx_idcx4 = self.findChild(QLabel, "tbx_idcx4")
        self.tbx_idcx5 = self.findChild(QLabel, "tbx_idcx5")
        self.tbx_idcx6 = self.findChild(QLabel, "tbx_idcx6")
        self.tbx_idcx7 = self.findChild(QLabel, "tbx_idcx7")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx8")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx9")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx10")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx11")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx12")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx13")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx14")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx15")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx16")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx17")
        self.tbx_idcx8 = self.findChild(QLabel, "tbx_idcx18")
        self.tbx_ppo = self.findChild(QLabel, "tbx_ppo")
        self.tbx_charges = self.findChild(QLabel, "tbx_charges")
        self.tbx_provider = self.findChild(QLabel, "tbx_provider")
        self.tbx_formtype = self.findChild(QLabel, "tbx_formtype")
        self.tbx_specialty = self.findChild(QLabel, "tbx_specialty")
        self.tbx_pricingstate = self.findChild(QLabel, "tbx_pricingstate")
        self.tbx_jurisdiction = self.findChild(QLabel, "tbx_jurisdiction")
        self.tbx_drg = self.findChild(QLabel, "tbx_drg")
        self.tbx_ratecode = self.findChild(QLabel, "tbx_ratecode")
        self.tbx_medicare = self.findChild(QLabel, "tbx_medicare")
        self.tbx_doi = self.findChild(QLabel, "tbx_doi")
        self.tbx_pdc = self.findChild(QLabel, "tbx_pdc")

        ## Timer
        self.lbl_timer = self.findChild(QLabel, "lbl_timer")

        ## Notification
        self.notif_icon_1 = self.findChild(QLabel, "notif_icon_1")
        self.notif_text_1 = self.findChild(QLabel, "notif_text_1")
        self.notif_icon_2 = self.findChild(QLabel, "notif_icon_2")
        self.notif_text_2 = self.findChild(QLabel, "notif_text_2")
        self.notif_icon_3 = self.findChild(QLabel, "notif_icon_3")
        self.notif_text_3 = self.findChild(QLabel, "notif_text_3")
        self.notif_icon_4 = self.findChild(QLabel, "notif_icon_4")
        self.notif_text_4 = self.findChild(QLabel, "notif_text_4")
        self.notif_icon_5 = self.findChild(QLabel, "notif_icon_5")
        self.notif_text_5 = self.findChild(QLabel, "notif_text_5")

        ## Tables
        self.tbl_services = self.findChild(QTableWidget, "tbl_services")
        self.tbl_opreport = self.findChild(QTableWidget, "tbl_opreport")
        self.tbl_encoderpro = self.findChild(QTableWidget, "tbl_encoderpro")
        self.tbl_duplicates = self.findChild(QTableWidget, "tbl_duplicates")
        self.tbl_aaos = self.findChild(QTableWidget, "tbl_aaos")

        ## Buttons
        self.btn_connect2web = self.findChild(QToolButton, "btn_connect2web")
        self.btn_info1 = self.findChild(QToolButton, "btn_info1")
        self.btn_pathway = self.findChild(QToolButton, "btn_pathway")
        self.btn_guidelines = self.findChild(QToolButton, "btn_guidelines")
        self.btn_documentation = self.findChild(QToolButton, "btn_documentation")
        self.btn_uploadallocation = self.findChild(QToolButton, "btn_uploadallocation")
        self.btn_opreport = self.findChild(QToolButton, "btn_opreport")
        self.btn_billimage = self.findChild(QToolButton, "btn_billimage")
        self.btn_aaos = self.findChild(QToolButton, "btn_aaos")
        self.btn_nextbill = self.findChild(QToolButton, "btn_nextbill")
        self.btn_settings = self.findChild(QToolButton, "btn_settings")
        self.btn_diagidcx_more = self.findChild(QToolButton, "btn_diagidcx_more")
        self.btn_diagidcx_close = self.findChild(QPushButton, "btn_diagidcx_close")

        ## Settings
        self.frm_settings = self.findChild(QFrame, "frm_settings")
        self.tbx_stratawareurl = self.findChild(QLineEdit, "tbx_stratawareurl")
        self.tbx_appscaling = self.findChild(QLineEdit, "tbx_appscaling")
        self.btn_settings_ok = self.findChild(QPushButton, "btn_settings_ok")
        self.btn_settings_cancel = self.findChild(QPushButton, "btn_settings_cancel")
        self.lbl_statusbar = self.findChild(QLabel, "lbl_statusbar")
        self.btn_settings_senderrorlogs = self.findChild(QPushButton, "btn_settings_senderrorlogs")
        self.btn_revalidate_version = self.findChild(QPushButton, "btn_revalidate_version")
        self.btn_restart_browser = self.findChild(QPushButton, "btn_restart_browser")
        self.version_edgeapp = self.findChild(QLabel, "version_edgeapp")
        self.version_webdriver = self.findChild(QLabel, "version_webdriver")
        self.force_appscaling = self.findChild(QCheckBox, "chk_override_scaling")
        
        ## Button functions
        self.btn_settings.clicked.connect(self.open_settings)
        self.btn_settings_cancel.clicked.connect(self.cancel_settings)
        self.btn_settings_ok.clicked.connect(self.save_settings)
        self.btn_settings_senderrorlogs.clicked.connect(self.send_errorlogs)
        self.btn_uploadallocation.clicked.connect(self.upload_allocation)
        self.btn_connect2web.clicked.connect(self.connect2web)
        self.btn_guidelines.clicked.connect(self.show_guidelines)
        self.btn_pathway.clicked.connect(self.display_pathway_rules)
        self.btn_info1.clicked.connect(self.show_info1)
        self.btn_nextbill.clicked.connect(self.next_allocation)
        self.btn_opreport.clicked.connect(self.ocr_reader)
        self.btn_inkling.clicked.connect(self.search_inkling)
        self.btn_billimage.clicked.connect(self.bill_image)
        self.btn_aaos.clicked.connect(self.upload_cpt_aaos)
        self.btn_revalidate_version.clicked.connect(self.revalidate_edge_version)
        self.btn_restart_browser.clicked.connect(self.restart_browser)
        self.btn_diagidcx_more.clicked.connect(lambda: self.frm_diagidcx.setVisible(True) if not self.frm_diagidcx.isVisible() else self.frm_diagidcx.setVisible(False))
        self.btn_diagidcx_close.clicked.connect(lambda: self.frm_diagidcx.setVisible(False))

        self.tbl_encoderpro.itemDoubleClicked.connect(self.tbl_encoderpro_onItemDoubleclicked)
        
        # Clean up before booting because you have default values in your textboxes/listboxes/labels
        self.clear_all_fields()
        self.revalidate_edge_version()

        # Load ASPT documents
        # self.show_guidelines()

        atexit.register(self.exit_cleanup)
        signal.signal(signal.SIGTERM, self.exit_cleanup)
        signal.signal(signal.SIGINT, self.exit_cleanup)
        sys.excepthook = self.excepthook_handler

        self.inkling_base_url = "https://libertymutual-sso.read.inkling.com/a/b/35d180377f0146678e78f6b60d743666"
        self.inkling_jurisdictions = [
            # State Abbv, State Full, EMG, Surgery, ASC, Facility Inpatient, Facility Outpatient
            # Users should still follow ASPT guidelines
            # By default, the tool would still extract CPT code information from EncoderPro (NCCI) 
            # even if the ASPT guideline states that it doesn't allow NCCI as coding authority
            ["AL", "Alabama", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["AK", "Alaska", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["AR", "Arkansas", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["AZ", "Arizona", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["CA", "California", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["CO", "Colorado", "AAOS", "AAOS", "NCCI", "NCCI", "NCCI"],
            ["CT", "Connecticut", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["DE", "Delaware", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["FL", "Florida", "AAOS", "AAOS", "NCCI", "NCCI", "NCCI"],
            ["GA", "Georgia", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["HI", "Hawaii", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["ID", "Idaho", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["IL", "Illinois", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["IN", "Indiana", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["IA", "Iowa", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["KS", "Kansas", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["KY", "Kentucky", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["LA", "Louisiana", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["ME", "Maine", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["MD", "Maryland", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["MA", "Massachusetts", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["MI", "Michigan", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["MN", "Minnesota", "AAOS", "AAOS", "NCCI", "NCCI", "NCCI"],
            ["MS", "Mississippi", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["MO", "Missouri", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["MT", "Montana", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["NE", "Nebraska", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["NV", "Nevada", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["NH", "New Hampshire", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["NJ", "New Jersey", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["NM", "New Mexico", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["NY", "New York", "AAOS", "AAOS", "NCCI", "NCCI", "NCCI"],
            ["NC", "North Carolina", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["ND", "North Dakota", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["OH", "Ohio", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["OK", "Oklahoma", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["OR", "Oregon", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["PA", "Pennsylvania", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["RI", "Rhode Island", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["SC", "South Carolina", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["SD", "South Dakota", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["TN", "Tennessee", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["TX", "Texas", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["UT", "Utah", "AAOS", "AAOS", "NCCI", "NCCI", "NCCI"],
            ["VT", "Vermont", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["VA", "Virginia", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["WA", "Washington", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["DC", "District of Columbia", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["WV", "West Virginia", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["WI", "Wisconsin", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"],
            ["WY", "Wyoming", "NCCI", "NCCI", "NCCI", "NCCI", "NCCI"]
        ]

        self.bc_charge_allowed = [
            "DE", "IN", "KS", "MT", "NY", "TX", "FL"
        ]

        # Default app settings
        self.tbx_appscaling.setText("1")
        self.qt_scale_factor = "1"
        self.setWindowTitle("Bill Review Assist")
        self.tbx_stratawareurl.setText(self.config["BRAT_CONFIG"]["url"])
        self.frm_settings.setVisible(False)
        self.frm_diagidcx.setVisible(False)

        QFontDatabase.addApplicationFont(f"{os.getcwd()}/resources/fonts/Aptos-Display.ttf")
        QFontDatabase.addApplicationFont(f"{os.getcwd()}/resources/fonts/Aptos-Black.ttf")
        QFontDatabase.addApplicationFont(f"{os.getcwd()}/resources/fonts/Aptos-SemiBold.ttf")
        QFontDatabase.addApplicationFont(f"{os.getcwd()}/resources/fonts/Aptos-ExtraBold.ttf")

        self.show()

    def excepthook_handler(self, type, value, trace):
        tb = traceback.extract_tb(trace)[0]

        fname = tb.filename
        line = tb.lineno
        fn_name = tb.name
        code_line = tb.line

        dt = datetime.datetime.now().strftime("%m/%d/%Y %I:%M %p")
        with open("error.log", "a") as log:
            log.write(f"Date/Time: {dt}\nFilename: {fname}\nFunction Name: {fn_name}\nLine #: {line}\nCode Line: {code_line}\n\nFull Traceback:\n{type}\n{value}\n{trace}\n----------\n")
        self.statusbar_update(f"The tool encountered an error during execution. Please try again.")
        # sys.exit()

    def exit_cleanup(self):
        try:
            self.driver.close()
            self.driver.quit()
        except:
            return
        
    def send_errorlogs(self):
        ol = win32.Dispatch("Outlook.Application")
        mail = ol.CreateItem(0)
        mail.Attachments.Add(f"{os.getcwd()}\\error.log")
        mail.Display()

    def show_notification(self, icon, notif_text, font_color, bgcolor):
        # Definitely think of something more elegant than this monstrosity
        if self.notif_text_1.text() == "":
            # 50, 210
            self.notif_icon_1.setText(icon)
            self.notif_icon_1.setStyleSheet(f"background-color: {bgcolor}; color: {font_color};")
            self.notif_text_1.setText(notif_text)
            self.notif_text_1.setStyleSheet("color: black;")
        elif self.notif_text_2.text() == "":
            # 260, 210
            self.notif_icon_2.setText(icon)
            self.notif_icon_2.setStyleSheet(f"background-color: {bgcolor}; color: {font_color};")
            self.notif_text_2.setText(notif_text)
            self.notif_text_2.setStyleSheet("color: black;")
        elif self.notif_text_3.text() == "":
            # 470, 210
            self.notif_icon_3.setText(icon)
            self.notif_icon_3.setStyleSheet(f"background-color: {bgcolor}; color: {font_color};")
            self.notif_text_3.setText(notif_text)
            self.notif_text_3.setStyleSheet("color: black;")
        elif self.notif_text_4.text() == "":
            # 50, 230
            self.notif_icon_4.setText(icon)
            self.notif_icon_4.setStyleSheet(f"background-color: {bgcolor}; color: {font_color};")
            self.notif_text_4.setText(notif_text)
            self.notif_text_4.setStyleSheet("color: black;")
        elif self.notif_text_5.text() == "":
            # 260, 230
            self.notif_icon_5.setText(icon)
            self.notif_icon_5.setStyleSheet(f"background-color: {bgcolor}; color: {font_color};")
            self.notif_text_5.setText(notif_text)
            self.notif_text_5.setStyleSheet("color: black;")

    def clear_all_fields(self):
        tmp_str = self.tbx_stratawareurl.text()

        for w in self.findChildren(QFrame):
            if w.objectName()[:3] == "lbl" or w.objectName()[:3] == "not":
                w.setText("")
                w.setStyleSheet("background-color: rgba(0,0,0,0);\nborder-color: rgba(0,0,0,0);")

            if w.objectName()[:3] == "tbl":
                w.setRowCount(0)

        for w in self.findChildren(QLabel):
            if w.objectName()[:3] == "tbx":
                w.setText("")
        
        self.lbl_errorpropensity.setStyleSheet("background-color: rgba(0,0,0,0); color: black; border: 1px; border-width: 1px; border-style: none; border-radius: 7px;")
        self.tbx_stratawareurl.setText(tmp_str)
        self.lbl_statusbar.setText("Idle")

        self.time = QTime(0, 0, 0)
        self.lbl_timer.setText("00:00:00")
        self.start_timer = False

    def open_settings(self):
        self.tmp_strata_url = self.tbx_stratawareurl.text()

        self.config.read("config.ini")
        if "BRAT_CONFIG" in self.config:
            self.tbx_stratawareurl.setText(self.config["BRAT_CONFIG"]["url"])
            self.force_appscaling.setChecked(bool(int(self.config["BRAT_CONFIG"]["app_scaling_override"])))
            self.tbx_appscaling.setText(self.config["BRAT_CONFIG"]["app_scaling_override_value"])

        self.frm_settings.setVisible(True)

    def cancel_settings(self):
        self.tbx_stratawareurl.setText(self.tmp_strata_url)
        self.frm_settings.setVisible(False)

    def save_settings(self):
        self.strataware_url = self.tbx_stratawareurl.text()
        
        if "BRAT_CONFIG" in self.config:
            self.config["BRAT_CONFIG"]["url"] = self.strataware_url
            self.config["BRAT_CONFIG"]["app_scaling_override"] = str(self.force_appscaling.checkState())
            self.config["BRAT_CONFIG"]["app_scaling_override_value"] = self.tbx_appscaling.text()

        with open("config.ini", "w") as f:
            self.config.write(f)

        if self.force_appscaling.isChecked():
            ctypes.windll.user32.MessageBoxW(0, "REMINDER: Restart app to apply new APP SCALING value.", "Bill Review Assist", 0)

        if not self.tmp_strata_url == self.strataware_url:
            restart_app = ctypes.windll.user32.MessageBoxW(0, "Changed Strataware URL. Do you want to restart the browser?", "Bill Review Assist", 4)
            if restart_app == 6:
                self.restart_browser()

        self.frm_settings.setVisible(False)

    def restart_browser(self):
        if not  self.driver == None:
            self.driver.close()
            self.driver.quit()
            self.driver = None
            time.sleep(3)

        self.connect2web()

    def window_focus(self, win_name):
        windows = self.driver.window_handles
        for window in windows:
            self.driver.switch_to.window(window)
            if win_name in self.driver.title:
                break
        time.sleep(3)

    def statusbar_update(self, text):
        self.lbl_statusbar.setText(text)
        QCoreApplication.processEvents()

    def ui_to_front(self):
        hwnd = self.winId()
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_SHOWWINDOW | win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    def revalidate_edge_version(self):
        try:
            file_info = win32api.GetFileVersionInfo(f"{self.documents_folder}\Selenium\msedgedriver.exe", "\\")
            ms = file_info["FileVersionMS"]
            ls = file_info["FileVersionLS"]
        
            self.edge_version = self.fso.GetFileVersion("c:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")
            self.webd_version = ".".join([str(win32api.HIWORD(ms)), str(win32api.LOWORD(ms)), str(win32api.HIWORD(ls)), str(win32api.LOWORD(ls))])

            self.version_edgeapp.setText(self.edge_version)
            self.version_webdriver.setText(self.webd_version)
            QCoreApplication.processEvents()
        except:
            self.version_webdriver.setText("0")

    def update_edge_browser(self, folder_path):
        if self.edge_version == self.webd_version:
            print("Versions match")
            return

        url = f"https://msedgedriver.microsoft.com/{self.edge_version}/edgedriver_win64.zip"

        if not os.path.isdir(f"{folder_path}\Selenium"):
            os.mkdir(f"{folder_path}\Selenium")

        if os.path.isdir(f"{folder_path}\Selenium\Driver_Notes"):
            rmtree(f"{folder_path}\Selenium\Driver_Notes")

        if os.path.exists(f"{folder_path}\Selenium\edgedriver_win64.zip"):
            os.remove(f"{folder_path}\Selenium\edgedriver_win64.zip")

        if os.path.exists(f"{folder_path}\Selenium\msedgedriver.exe"):
            os.system("taskkill /F /IM msedgedriver.exe")
            time.sleep(3)
            os.remove(f"{folder_path}\Selenium\msedgedriver.exe")

        download(url, f"{folder_path}\Selenium")
        with ZipFile(f"{folder_path}\Selenium\edgedriver_win64.zip") as zf:
            zf.extractall(path=f"{folder_path}\Selenium")
            zf.close()

    def connect2web(self):
        if self.driver == None:
            # Comment out user-data-dir if you are having issues with browser session
            username = getpass.getuser()
            user_data_dir = f'C:\\Users\\{username}\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default'

            self.update_edge_browser(folder_path=self.documents_folder)

            service = Service(executable_path=f"{self.documents_folder}\Selenium\msedgedriver.exe")
            options = webdriver.EdgeOptions()
            options.add_experimental_option("useAutomationExtension", False)
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_argument("--disable-cache")
            options.add_argument(f"--user-data-dir={user_data_dir}")

            # Added as a workaround to fix login errors in EncoderPro
            # Not sure if this actually fixes it though. lmao
            options.add_argument(f"--remote-debugging-port={self.port}")
            print(f"Opened port number {self.port}")

            self.statusbar_update("Opening Strataware...")
            self.driver = webdriver.Edge(service=service, options=options)
            self.driver.implicitly_wait(15)
            self.driver.get(self.strataware_url)
            
            self.driver.implicitly_wait(10) # Added because I had page timeout issues in Strataware. Feel free to reduce/optimize.
            time.sleep(5)
            while self.driver.title == "Strataware Maintenance Page":
                time.sleep(3)

            self.statusbar_update("Opening SmartCoder...")
            strata_url = self.driver.current_url.split("/")[2]
            self.driver.get(f"https://{strata_url}/strataware/strataware/smartcoder/mainframe.aspx?Screen=SMARTCODER")
            self.driver.execute_script("window.open('https://www.encoderpro.com')")
            self.driver.execute_script(f"window.open('https://libertymutual-sso.read.inkling.com/a/b/35d180377f0146678e78f6b60d743666/p/bc126e50ae9e46508598b41bd5d667bf')")

            self.statusbar_update("Ready. Make sure you are logged in EncoderPro before proceeding.")

    def search_inkling(self):
        self.window_focus("Liberty Mutual")
        self.ui_to_front()

        state_jurisdiction = self.tbx_jurisdiction.text()
        for jurisdiction in self.inkling_jurisdictions:
            if jurisdiction[0] == state_jurisdiction:
                state_jurisdiction = jurisdiction[1]
                break

        list_items = self.driver.find_elements(By.XPATH, "//li[@class='chapter ']")
        for li in list_items:
            if "State Resources Protocols" in li.get_attribute("innerText"):
                if li.get_attribute("class") == "chapter selected-chapter":
                    break

                li.click()
                break

        WebDriverWait(self.driver, 30).until(EC.visibility_of_any_elements_located((By.XPATH, "//span[@class='title']")))
        spans = self.driver.find_elements(By.XPATH, "//span[@class='title']")
        for span in spans:
            if f"{state_jurisdiction} State Handling" in span.get_attribute("innerText"):
                span.click()
                break
        
    def element_exists(self, xpath):
        try:
            self.driver.find_element(By.XPATH, xpath)
        except NoSuchElementException:
            return False
        
        return True

    def extract_strataware(self):
        def find_xpath(element, element_name):
            return self.driver.find_element(By.XPATH, f"//{element}[@id='{element_name}']")

        self.pathway_rules = []
        self.services_array = []
        self.cpt_array = []

        self.window_focus("SmartCoder")
        self.ui_to_front()
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame('applFrame')
        self.driver.find_element(By.XPATH, "//input[@id='billCtrlNumber']").clear()
        self.driver.find_element(By.XPATH, "//input[@id='billCtrlNumber']").send_keys(self.lbl_icn.text())
        self.driver.find_element(By.XPATH, "//input[@id='billCtrlNumber']").send_keys(Keys.TAB)

        # Wait for the demographics fields to load properly
        # Optimization opportunity +1
        self.statusbar_update("Extracting data from Strataware...")
        while self.driver.find_element(By.XPATH, "//input[@id='provProviderType']").get_attribute("value") == "":
            time.sleep(1)

        # Strataware > Demographics
        self.tbx_claimant.setText(find_xpath("div", "billClaimDescName").get_attribute("innerText"))
        self.tbx_provider.setText(find_xpath("input", "provProviderType").get_attribute("value"))
        self.tbx_formtype.setText(find_xpath("input", "formType").get_attribute("value"))
        self.tbx_specialty.setText(find_xpath("input", "provSpecialty").get_attribute("value"))
        self.tbx_idcx1.setText(find_xpath("input", "clientDiagICDxNum1").get_attribute("value"))
        self.tbx_idcx2.setText(find_xpath("input", "clientDiagICDxNum2").get_attribute("value"))
        self.tbx_idcx3.setText(find_xpath("input", "clientDiagICDxNum3").get_attribute("value"))
        self.tbx_idcx4.setText(find_xpath("input", "clientDiagICDxNum4").get_attribute("value"))
        self.tbx_idcx5.setText(find_xpath("input", "clientDiagICDxNum5").get_attribute("value"))
        self.tbx_idcx6.setText(find_xpath("input", "clientDiagICDxNum6").get_attribute("value"))
        self.tbx_idcx7.setText(find_xpath("input", "clientDiagICDxNum7").get_attribute("value"))
        self.tbx_idcx8.setText(find_xpath("input", "clientDiagICDxNum8").get_attribute("value"))
        self.tbx_drg.setText(find_xpath("input", "provDRG").get_attribute("value"))
        self.tbx_ratecode.setText(find_xpath("input", "rateCode").get_attribute("value"))
        self.tbx_medicare.setText(find_xpath("input", "provMedicare").get_attribute("value"))
        self.tbx_doi.setText(find_xpath("div", "billClaimDescDOI").get_attribute("innerText"))
        self.tbx_pdc.setText(find_xpath("div", "billPaymentDirectionCode").get_attribute("innerText"))

        # More IDCX codes
        find_xpath("img", "imgDiagDetailsBtn").click()
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame(2)
        self.tbx_idcx9.setText(find_xpath("input", "diagCode9").get_attribute("value"))
        self.tbx_idcx10.setText(find_xpath("input", "diagCode10").get_attribute("value"))
        self.tbx_idcx11.setText(find_xpath("input", "diagCode11").get_attribute("value"))
        self.tbx_idcx12.setText(find_xpath("input", "diagCode12").get_attribute("value"))
        self.tbx_idcx13.setText(find_xpath("input", "diagCode13").get_attribute("value"))
        self.tbx_idcx14.setText(find_xpath("input", "diagCode14").get_attribute("value"))
        self.tbx_idcx15.setText(find_xpath("input", "diagCode15").get_attribute("value"))
        self.tbx_idcx16.setText(find_xpath("input", "diagCode16").get_attribute("value"))
        self.tbx_idcx17.setText(find_xpath("input", "diagCode17").get_attribute("value"))
        self.tbx_idcx18.setText(find_xpath("input", "diagCode18").get_attribute("value"))
        self.driver.execute_script("javascript: OnOK()")

        # Strataware > Services
        time.sleep(3)
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame('menuFrame')
        self.driver.execute_script("javascript:onMenuServices()")

        # menuServices
        time.sleep(3)
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame('applFrame')
        self.tbx_charges.setText(find_xpath("div", "extSvcRunningTotal").get_attribute("innerText"))
        find_xpath("span", "tabName0").click() # Services > Bill -- no shortcut key for this one
        self.tbx_ppo.setText(find_xpath("div", "BillPrePPO").get_attribute("innerText"))

        if not self.tbx_charges.text().replace("$", "").replace(",", "") == self.lbl_totalcharges.text().replace("$", "").replace(",", ""):
            self.show_notification("!", "CHRGS MISMATCH", "white", "blue")

        # This is not available in the test environment.
        # self.driver.find_element(By.XPATH, "//html/body[@id='SmartCoder']/label[@id='capBillInfo']/b/span/u").click()

        # Extract Services/CPT/Modifiers table
        self.driver.switch_to.frame("iframeServiceGrid")
        elem = self.driver.find_element(By.NAME, "sldDOSRow0Col1")
        elem_id = None
        x = 0

        while elem is not None:
            if elem.tag_name == "input":
                temp = [
                    find_xpath("input", f"sldDOSRow{x}Col1").get_attribute("value"),
                    find_xpath("input", f"sldPOSRow{x}Col2").get_attribute("value"),
                    find_xpath("input", f"sldTOSRow{x}Col3").get_attribute("value"),
                    find_xpath("input", f"sldServiceRow{x}Col4").get_attribute("value"),
                    find_xpath("input", f"sldServiceRow{x}Col5").get_attribute("value"),
                    find_xpath("input", f"sldM1Row{x}Col6").get_attribute("value"),
                    find_xpath("input", f"sldM2Row{x}Col7").get_attribute("value"),
                    find_xpath("input", f"sldM3Row{x}Col8").get_attribute("value"),
                    find_xpath("input", f"sldM4Row{x}Col9").get_attribute("value"),
                    find_xpath("input", f"sldQtyRow{x}Col10").get_attribute("value"),
                    find_xpath("input", f"slddblChargeRow{x}Col11").get_attribute("value"),
                    find_xpath("input", f"slddblAllowanceRow{x}Col12").get_attribute("value"),
                    find_xpath("input", f"sldMsgRow{x}Col13").get_attribute("value")
                ]
            else:
                temp = [
                    find_xpath("*", f"sldDOSRow{x}Col1").get_attribute("innerText"),
                    find_xpath("*", f"sldPOSRow{x}Col2").get_attribute("innerText"),
                    find_xpath("*", f"sldTOSRow{x}Col3").get_attribute("innerText"),
                    find_xpath("*", f"sldServiceRow{x}Col4").get_attribute("innerText"),
                    find_xpath("*", f"sldServiceRow{x}Col5").get_attribute("innerText"),
                    find_xpath("*", f"sldM1Row{x}Col6").get_attribute("innerText"),
                    find_xpath("*", f"sldM2Row{x}Col7").get_attribute("innerText"),
                    find_xpath("*", f"sldM3Row{x}Col8").get_attribute("innerText"),
                    find_xpath("*", f"sldM4Row{x}Col9").get_attribute("innerText"),
                    find_xpath("*", f"sldQtyRow{x}Col10").get_attribute("innerText"),
                    find_xpath("*", f"slddblChargeRow{x}Col11").get_attribute("innerText"),
                    find_xpath("*", f"slddblAllowanceRow{x}Col12").get_attribute("innerText"),
                    find_xpath("*", f"sldMsgRow{x}Col13").get_attribute("innerText")
                ]
            
            self.services_array.append(temp)

            if temp[3] != "": self.cpt_array.append(temp[3])
            if temp[4] != "": self.cpt_array.append(temp[4])

            elem = None
            x = x + 1
            elem_id = f"sldDOSRow{x}Col1"

            try:
                elem = self.driver.find_element(By.NAME, elem_id)
            except NoSuchElementException:
                elem = None

        # Services table prep
        self.tbl_services.clearContents()
        self.tbl_services.setRowCount(len(self.services_array))
        self.tbl_services.verticalScrollBar().setSingleStep(7)
        self.tbl_services.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.tbl_services.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.tbl_services.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(10, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(11, QHeaderView.ResizeToContents)
        self.tbl_services.horizontalHeader().setSectionResizeMode(12, QHeaderView.ResizeToContents)
        self.tbl_services.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # AAOS table prep
        self.tbl_aaos.clearContents()
        self.tbl_aaos.setRowCount(0)
        self.tbl_aaos.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tbl_aaos.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)

        for x, arr in enumerate(self.services_array[:-1]):
            self.tbl_services.setItem(x, 0, QTableWidgetItem(arr[0]))
            self.tbl_services.item(x, 0).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 1, QTableWidgetItem(arr[1]))
            self.tbl_services.item(x, 1).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 2, QTableWidgetItem(arr[2]))
            self.tbl_services.item(x, 2).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 3, QTableWidgetItem(arr[3]))
            self.tbl_services.item(x, 3).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.item(x, 3).setToolTip(arr[3])
            self.tbl_services.setItem(x, 4, QTableWidgetItem(arr[4]))
            self.tbl_services.item(x, 4).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.item(x, 4).setToolTip(arr[4])
            self.tbl_services.setItem(x, 5, QTableWidgetItem(arr[5]))
            self.tbl_services.item(x, 5).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 6, QTableWidgetItem(arr[6]))
            self.tbl_services.item(x, 6).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 7, QTableWidgetItem(arr[7]))
            self.tbl_services.item(x, 7).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 8, QTableWidgetItem(arr[8]))
            self.tbl_services.item(x, 8).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 9, QTableWidgetItem(arr[9]))
            self.tbl_services.item(x, 9).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 10, QTableWidgetItem(arr[10]))
            self.tbl_services.item(x, 10).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 11, QTableWidgetItem(arr[11]))
            self.tbl_services.item(x, 11).setTextAlignment(Qt.AlignCenter)
            self.tbl_services.setItem(x, 12, QTableWidgetItem(arr[12]))
            self.tbl_services.item(x, 12).setTextAlignment(Qt.AlignCenter)

        # Strataware > Services > Pathway
        ActionChains(self.driver).key_down(Keys.ALT).send_keys("w").perform()
        ActionChains(self.driver).reset_actions()
        time.sleep(3)
        self.window_focus("Bill Rules")
        self.ui_to_front()
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame("frmBillRules")
        
        x = 1
        while True:
            # Loop Bill Rules contents since this contains vital information
            # regarding bill review process per claim
            try:
                # This is pretty straight-forward, yes?
                # No? Fine... This is just a check if there is an element
                # with ID lblRuleIDX in the Bill Rules window (X = Number, duh)
                elem = find_xpath("span", f"lblRuleID{x}")
            except NoSuchElementException:
                break # Didn't find any elements with that ID

            temp = [
                # Pass RuleID, Name, and Description to a temp array.
                # To be used when displaying Bill Rules in the app
                find_xpath("span", f"lblRuleID{x}").get_attribute("innerText"),
                find_xpath("span", f"lblRuleName{x}").get_attribute("innerText"),
                find_xpath("span", f"lblRuleDescription{x}").get_attribute("innerText")
            ]

            self.pathway_rules.append(temp)
            elem = None # Clear the elem variable
            x = x + 1

        self.display_pathway_rules()

        # Might have a few tweaks/optimizations to be done here
        # This part re-focuses SmartCoder window, which for some reason is not being focused
        # properly by Python... So if you find any other alternatives to re-focus SmartCoder, the better.
        self.window_focus('SmartCoder')
        self.ui_to_front()
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame('applFrame')
        self.driver.switch_to.frame('iframeServiceGrid')
        find_xpath("img", "sldBRLogImgRow0Col15BRLogImg").click() # Click Bill Review Logs icon
        time.sleep(3)

        # The Bill Review Log pop up window
        # Get Pricing and Jurisdiction states from Strataware > Services > Bill Review Logs
        self.window_focus('Bill Review Log')
        self.ui_to_front()
        spans = find_xpath("ul", "tree").find_elements(By.TAG_NAME, "span")
        for index, span in enumerate(spans):
            # Loop all elements inside the pop up window and look for the keywords
            if "pricing state" in span.get_attribute("innerText").lower():
                i = index + 2
                self.tbx_pricingstate.setText(self.driver.find_element(By.XPATH, f"(//span)[{i}]").get_attribute("innerText"))

            if "jurisdiction state" in span.get_attribute("innerText").lower():
                i = index + 2
                self.tbx_jurisdiction.setText(self.driver.find_element(By.XPATH, f"(//span)[{i}]").get_attribute("innerText"))

        if not self.tbx_jurisdiction.text().lower() == self.lbl_jurisdiction.text().lower():
            self.show_notification("!", "JURIS MISMATCH", "white", "blue")

        for state in self.bc_charge_allowed:
            if self.tbx_pricingstate.text() == state:
                self.show_notification("!", "GREATER THAN BC STATE", "white", "blue")
                break

        for state in self.inkling_jurisdictions:
            # EMG, Surgery, ASC, Facility
            if self.tbx_pricingstate.text() == state[0]:
                if "emg" in self.lbl_holdgroup.text().lower():
                    if state[2] == "AAOS":
                        self.show_notification("!", "CHECK AAOS TAB", "white", "blue")
                        time_start = datetime.datetime.now()
                        for cpt in self.cpt_array:
                            self.statusbar_update("Checking AAOS data. Please wait... Time Elapsed: {}".format(datetime.datetime.now() - time_start))
                            if len(cpt) == 5: self.extract_aaos(cpt)
                if "surg" in self.lbl_holdgroup.text().lower():
                    if state[3] == "AAOS":
                        self.show_notification("!", "CHECK AAOS TAB", "white", "blue")
                        time_start = datetime.datetime.now()
                        for cpt in self.cpt_array:
                            self.statusbar_update("Checking AAOS data. Please wait... Time Elapsed: {}".format(datetime.datetime.now() - time_start))
                            if len(cpt) == 5: self.extract_aaos(cpt)
                if "faci" in self.lbl_holdgroup.text().lower():
                    if state[5] == "AAOS" or state[6] == "AAOS":
                        self.show_notification("!", "CHECK AAOS TAB", "white", "blue")
                        time_start = datetime.datetime.now()
                        for cpt in self.cpt_array:
                            self.statusbar_update("Checking AAOS data. Please wait... Time Elapsed: {}".format(datetime.datetime.now() - time_start))
                            if len(cpt) == 5: self.extract_aaos(cpt)

        self.window_focus('SmartCoder')
        self.ui_to_front()
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame('applFrame')
        find_xpath("span", "tabName3").click() # Click Strataware > Services > Duplicates
        time.sleep(5)

        self.driver.switch_to.default_content()
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        self.statusbar_update("Checking for duplicates...")

        for iframe in iframes:
            if 'ERRORV2' in iframe.get_attribute('src'):
                self.driver.switch_to.frame(self.driver.find_element(By.XPATH, "//html/body/iframe"))
                self.driver.find_element(By.ID, 'btnImg').click()
                break
            else:
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame(self.driver.find_element(By.ID, 'applFrame'))
                self.driver.switch_to.frame(self.driver.find_element(By.ID, 'iframeDuplicateGrid'))
                i = 0

                WebDriverWait(self.driver, 30).until(EC.visibility_of_any_elements_located((By.XPATH, "//table[@id='tblDuplicateDetail']")))
                dup_table = find_xpath("table", "tblDuplicateDetail").find_element(By.TAG_NAME, "tbody")
                self.dup_array = []

                for table_data in dup_table.find_elements(By.XPATH, "./tr"):
                    status_dup = table_data.find_element(By.XPATH, ".//input[@id='dupState']")
                    if status_dup.get_attribute("value") == "Exact Dup.":
                        if self.element_exists("//input[@id='txtDuplicateList']"):
                            select_dup = table_data.find_element(By.XPATH, "//input[@id='txtDuplicateList']")
                            duplicate_text = select_dup.get_attribute("value")
                        else:
                            select_dup = Select(table_data.find_element(By.XPATH, f".//select[@id='lstDuplicateList{i}']"))
                            duplicate_text = select_dup.first_selected_option.text                       
                    else:
                        select_dup = table_data.find_element(By.XPATH, f".//input[@id='txtDuplicateList']")
                        duplicate_text = select_dup.get_attribute("value")
                    select_act = Select(table_data.find_element(By.XPATH, ".//select[@id='lstDuplicateAction']"))

                    temp_array = [
                        duplicate_text,
                        select_act.first_selected_option.text,
                        status_dup.get_attribute("value"),
                        table_data.find_element(By.XPATH, "(.//input[@id='txtDuplicateDOS'])[1]").get_attribute("value"),
                        table_data.find_element(By.XPATH, ".//input[@id='txtDuplicatePOS']").get_attribute("value"),
                        table_data.find_element(By.XPATH, "(.//input[@id='txtDuplicateDOS'])[2]").get_attribute("value"),
                        table_data.find_element(By.XPATH, ".//input[@id='txtDuplicateServiceCode']").get_attribute("value"),
                        table_data.find_element(By.XPATH, ".//input[@id='txtDuplicateM1']").get_attribute("value"),
                        table_data.find_element(By.XPATH, ".//input[@id='txtDuplicateM2']").get_attribute("value"),
                        table_data.find_element(By.XPATH, ".//input[@id='txtDuplicateM3']").get_attribute("value"),
                        table_data.find_element(By.XPATH, ".//input[@id='txtDuplicateM4']").get_attribute("value"),
                        table_data.find_element(By.XPATH, ".//input[@id='txtDuplicateQuantity']").get_attribute("value"),
                        table_data.find_element(By.XPATH, ".//input[@id='txtDuplicateAllowance']").get_attribute("value")
                    ]
                    self.dup_array.append(temp_array)
                    i = i + 1

                break
        
        dup_found = False
        for ind_array in self.dup_array:
            if "Exact Dup." in ind_array:
                dup_found = True
                break

        if dup_found:
            self.show_notification("X", "FOUND DUPLICATE(S)", "white", "red")
        else:
            self.show_notification("!", "NO DUPLICATE(S) FOUND", "white", "blue")

        self.display_duplicates()
        self.extract_encoderpro()
        self.extract_cptconflicts()

        self.start_timer = True

    def display_duplicates(self):
        if len(self.dup_array) > 0:
            self.tbl_duplicates.setRowCount(len(self.dup_array))
        else:
            self.tbl_duplicates.setRowCount(0)
        
        self.tbl_duplicates.clearContents()
        self.tbl_duplicates.verticalHeader().setVisible(False)
        self.tbl_duplicates.verticalScrollBar().setSingleStep(7)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(10, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(11, QHeaderView.ResizeToContents)
        self.tbl_duplicates.horizontalHeader().setSectionResizeMode(12, QHeaderView.ResizeToContents)
        self.tbl_duplicates.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for x, arr in enumerate(self.dup_array):
            self.tbl_duplicates.setItem(x, 0, QTableWidgetItem(arr[0]))
            self.tbl_duplicates.item(x, 0).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 1, QTableWidgetItem(arr[1]))
            self.tbl_duplicates.item(x, 1).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 2, QTableWidgetItem(arr[2]))
            self.tbl_duplicates.item(x, 2).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 3, QTableWidgetItem(arr[3]))
            self.tbl_duplicates.item(x, 3).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 4, QTableWidgetItem(arr[4]))
            self.tbl_duplicates.item(x, 4).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 5, QTableWidgetItem(arr[5]))
            self.tbl_duplicates.item(x, 5).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 6, QTableWidgetItem(arr[6]))
            self.tbl_duplicates.item(x, 6).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 7, QTableWidgetItem(arr[7]))
            self.tbl_duplicates.item(x, 7).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 8, QTableWidgetItem(arr[8]))
            self.tbl_duplicates.item(x, 8).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 9, QTableWidgetItem(arr[9]))
            self.tbl_duplicates.item(x, 9).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 10, QTableWidgetItem(arr[10]))
            self.tbl_duplicates.item(x, 10).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 11, QTableWidgetItem(arr[11]))
            self.tbl_duplicates.item(x, 11).setTextAlignment(Qt.AlignCenter)
            self.tbl_duplicates.setItem(x, 12, QTableWidgetItem(arr[12]))
            self.tbl_duplicates.item(x, 12).setTextAlignment(Qt.AlignCenter)

        self.tbl_duplicates.resizeRowsToContents()
        self.show()
        QCoreApplication.processEvents()
    
    def display_pathway_rules(self):
        self.statusbar_update("Pathway rules displayed")

        if self.pathway_window == None:
            self.pathway_window = PathwayWindow()

        if len(self.pathway_rules) > 0:
            self.pathway_window.pathway_table.setRowCount(len(self.pathway_rules))
        else:
            self.pathway_window.pathway_table.setRowCount(0)

        self.pathway_window.pathway_table.clearContents()
        self.pathway_window.pathway_table.setColumnCount(3) 
        self.pathway_window.pathway_table.setHorizontalHeaderLabels(["RULE ID", "NAME", "DESCRIPTION"])
        self.pathway_window.pathway_table.verticalHeader().setVisible(False)
        self.pathway_window.pathway_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.pathway_window.pathway_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.pathway_window.pathway_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.pathway_window.pathway_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for x, arr in enumerate(self.pathway_rules):
            self.pathway_window.pathway_table.setItem(x, 0, QTableWidgetItem(arr[0]))
            self.pathway_window.pathway_table.item(x, 0).setTextAlignment(Qt.AlignCenter)
            self.pathway_window.pathway_table.setItem(x, 1, QTableWidgetItem(arr[1]))
            self.pathway_window.pathway_table.item(x, 1).setTextAlignment(Qt.AlignCenter)
            self.pathway_window.pathway_table.setItem(x, 2, QTableWidgetItem(arr[2]))
            self.pathway_window.pathway_table.item(x, 2).setTextAlignment(Qt.AlignCenter)

        self.pathway_window.pathway_table.resizeRowsToContents()
        self.pathway_window.show()
        QCoreApplication.processEvents()

    def extract_encoderpro(self):
        def find_xpath(element, element_name):
            return self.driver.find_element(By.XPATH, f"//{element}[@id='{element_name}']")
        
        self.statusbar_update("Extracting data from EncoderPro...")
        self.window_focus("EncoderPro.com")
        self.ui_to_front()

        # self.cpt_array = []
        cpt_codes = []
        mue_units = []
        tmp_object = []
        self.cpt_details = []

        if len(self.cpt_array) == 0:
            return

        for cpt in self.cpt_array:
            if len(cpt) == 5:
                # Filter CPT codes. Valid CPT code usually consists of 5 characters
                cpt_codes.append(cpt)

        cpt_codes = list(set(cpt_codes))
        cpt_codes.sort()
        
        # cpt = ", ".join(cpt_codes) # Join CPT codes separated by comma(s) for batch search in EncoderPro
        # Can't do batch search anymore, you fckn dumbass

        i = 0
        # Prep EncoderPro table in the UI
        self.tbl_encoderpro.clearContents()
        self.tbl_encoderpro.setRowCount(0)
        self.tbl_encoderpro.setRowCount(len(cpt_codes))
        self.tbl_encoderpro.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tbl_encoderpro.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tbl_encoderpro.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.tbl_encoderpro.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.tbl_encoderpro.horizontalHeader().setSectionResizeMode(4, QHeaderView.Fixed)
        self.tbl_encoderpro.setColumnWidth(4, 180)
        self.tbl_encoderpro.horizontalHeader().setSectionResizeMode(5, QHeaderView.Fixed)
        self.tbl_encoderpro.setColumnWidth(5, 80)
        self.tbl_encoderpro.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.tbl_encoderpro.verticalScrollBar().setSingleStep(9)
        self.tbl_encoderpro.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for cpt in cpt_codes:
            self.driver.get("https://www.encoderpro.com/epro/index.jsp")

            WebDriverWait(self.driver, 30).until(EC.visibility_of_any_elements_located((By.XPATH, "//input[@id='searchTermsMain']")))
            find_xpath("input", "searchTermsMain").send_keys(cpt)
            find_xpath("input", "searchTermsMain").send_keys(Keys.ENTER)

            if len(self.driver.find_elements(By.ID, 'LIST_BOX_TABLE_ID')) > 0:
                table = find_xpath("table", "LIST_BOX_TABLE_ID").find_element(By.TAG_NAME, "tbody")
                x = 0

                # Should throw an exception error if no CPT found
                for tr in table.find_elements(By.XPATH, "./tr"):
                    tmp_object = []

                    try:
                        cpt_code = tr.find_elements(By.TAG_NAME, "td")[1].get_attribute("innerText")
                        if cpt == cpt_code:
                            tmp_object.append(cpt_code)
                            tmp_object.append(tr.find_elements(By.TAG_NAME, "td")[2].get_attribute("innerText"))

                            if "hcpcs" in self.driver.current_url.lower():
                                self.driver.get(f"https://www.encoderpro.com/epro/hcpcsHandler.do?_k=102*{cpt}&_a=viewDetail")
                                WebDriverWait(self.driver, 30).until(EC.visibility_of_any_elements_located((By.XPATH, "//table[@id='code_lay']")))
                                tmp_object.append(self.driver.find_element(By.XPATH, "//table[@id='code_lay']").get_attribute("innerText"))
                                tmp_object.append("hcpcs")
                            else:
                                self.driver.get(f"https://www.encoderpro.com/epro/cptHandler.do?_k=101*{cpt}&_a=viewDetail")
                                WebDriverWait(self.driver, 30).until(EC.visibility_of_any_elements_located((By.XPATH, "//table[@id='code_lay']")))
                                tmp_object.append(self.driver.find_element(By.XPATH, "//table[@id='code_lay']").get_attribute("innerText"))
                                tmp_object.append("cpt")

                            infobox_td = self.driver.find_elements(By.XPATH, "//td[@class='infobox']")
                            pmt_found = False
                            ast_found = False

                            for td in infobox_td:
                                if "payment indicator" in td.get_attribute('innerText').lower():
                                    tbl1 = td.find_element(By.XPATH, "./table/tbody/tr/td")
                                    tbl2 = tbl1.find_element(By.XPATH, "./table/tbody")

                                    for tr in tbl2.find_elements(By.XPATH, "./tr"):
                                        if "payment indicator" in tr.get_attribute("innerText").lower():
                                            pmt_found = True
                                            pmt_ind = tr.get_attribute("innerText")
                                        else:
                                            if pmt_found:
                                                pmt_ind = f"{pmt_ind}\n{tr.get_attribute('innerText')}"
                                                break
                                    break

                            for td in infobox_td:
                                if "assist-at-surgery" in td.get_attribute("innerText").lower():
                                    tbl1 = td.find_element(By.XPATH, "./table/tbody/tr/td")
                                    tbl2 = tbl1.find_element(By.XPATH, "./table/tbody")

                                    for tr in tbl2.find_elements(By.XPATH, "./tr"):
                                        if "assist-at-surgery" in tr.get_attribute("innerText").lower():
                                            ast_found = True
                                            ast_ind = tr.get_attribute("innerText")
                                            break
                                    break

                            if pmt_found:
                                tmp_object.append(pmt_ind)
                            else:
                                tmp_object.append("N/A")

                            if ast_found:
                                tmp_object.append(ast_ind)
                            else:
                                tmp_object.append("N/A")

                            self.driver.get(f"https://www.encoderpro.com/epro/mueHandler.do?_k=6110*{cpt}&_a=view")
                            WebDriverWait(self.driver, 30).until(EC.visibility_of_any_elements_located((By.ID, "LIST_BOX_TABLE_ID")))
                            mue_units = self.driver.find_element(By.ID, 'LIST_BOX_TABLE_ID').find_element(By.TAG_NAME, 'tbody').find_element(By.TAG_NAME, 'tr').find_elements(By.TAG_NAME, 'td')[0].text

                            if not mue_units == "":
                                tmp_object.append(mue_units)
                            else:
                                tmp_object.append("N/A")

                            self.cpt_details.append(tmp_object)

                            self.tbl_encoderpro.setItem(i, 0, QTableWidgetItem(self.cpt_details[i][0]))
                            self.tbl_encoderpro.setItem(i, 1, QTableWidgetItem(self.cpt_details[i][1]))
                            self.tbl_encoderpro.setItem(i, 2, QTableWidgetItem(self.cpt_details[i][2]))
                            self.tbl_encoderpro.verticalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)
                            self.tbl_encoderpro.setItem(i, 3, QTableWidgetItem(self.cpt_details[i][6]))
                            self.tbl_encoderpro.item(i, 3).setTextAlignment(Qt.AlignCenter)

                            if ast_found:
                                pmt_ast = f"{self.cpt_details[i][4]}\n\nMod. 80 -{self.cpt_details[i][5]}"
                                self.tbl_encoderpro.setItem(i, 4, QTableWidgetItem(pmt_ast))
                            else:
                                self.tbl_encoderpro.setItem(i, 4, QTableWidgetItem(f"{self.cpt_details[i][4]}"))

                            i = i + 1
                            break
                    except:
                        continue
        
        QCoreApplication.processEvents()

    def extract_cptconflicts(self):
        if len(self.cpt_details) == 0:
            return

        self.statusbar_update("Extracting CPT conflicts...")
        self.window_focus("EncoderPro.com")
        self.ui_to_front()
        self.driver.get("https://www.encoderpro.com/epro/index.jsp")
        self.driver.find_element(By.XPATH, "//img[@id='morebtn']").click()

        time.sleep(3) # Added pause timer because there were instances were the page loaded too quickly before the driver triggers the wait function
        WebDriverWait(self.driver, 30).until(EC.visibility_of_any_elements_located((By.XPATH, "//input[@id='codes2Find0']")))

        i = 0
        for cpt in self.cpt_details:
            self.driver.find_element(By.XPATH, f"//input[@id='codes2Find{i}']").send_keys(cpt[0])
            i = i + 1

        self.driver.find_element(By.XPATH, "//input[@id='cciCheckId']").click()
        WebDriverWait(self.driver, 30).until(EC.visibility_of_any_elements_located((By.XPATH, "//table[@id='LIST_BOX_TABLE_ID']")))
        tbl_conflicts = self.driver.find_elements(By.XPATH, "//table[@id='LIST_BOX_TABLE_ID']/tbody")
        time.sleep(3)

        for tbl in tbl_conflicts:
            tmp_arr = []
            for tr in tbl.find_elements(By.XPATH, "./tr"):
                tmp_arr.append(tr.find_elements(By.TAG_NAME, "td")[0].get_attribute("innerText"))

            if len(tmp_arr) <= 1: tmp_arr.append("No conflicts")
            self.cpt_conflicts.append(tmp_arr)

        for r in range(self.tbl_encoderpro.rowCount()):
            ins_text = ""
            for cpt in self.cpt_conflicts:
                try:
                    if cpt[0] == self.tbl_encoderpro.item(r, 0).text():
                        if ins_text == "":
                            ins_text = cpt[1]
                        else:
                            ins_text = f"{ins_text} {cpt[1]}"

                        self.tbl_encoderpro.setItem(r, 5, QTableWidgetItem(" ".join(sorted(set(ins_text.split()), key=ins_text.index))))

                    if ins_text == "":
                        self.tbl_encoderpro.setItem(r, 5, QTableWidgetItem("Secondary code. Check NCCI if allowed to override."))
                except:
                    continue

        self.statusbar_update("Done")

    def tbl_encoderpro_onItemDoubleclicked(self, item):
        for cpt in self.cpt_details:
            if cpt[0] == self.tbl_encoderpro.item(item.row(), 0).text():
                if cpt[3] == "cpt":
                    self.driver.get(f"https://www.encoderpro.com/epro/modifiersHandler.do?_k=101*{cpt[0]}&_a=list")
                    break
                else:
                    self.driver.get(f"https://www.encoderpro.com/epro/modifiersHandler.do?_k=102*{cpt[0]}&_a=list")
                    break

        self.window_focus("EncoderPro.com")

    def error_propensity(self, billholdgroup, claimjurisdiction, riskstatus, billingtype, totalcharges, totalallowance):
        # Load the model and encoder
        clf = joblib.load('resources\\error_propensity\\model.pkl')
        encoder = joblib.load('resources\\error_propensity\\encoder.pkl')

        input_data = pd.DataFrame([[billholdgroup, claimjurisdiction, riskstatus, billingtype, totalcharges, totalallowance]], 
                                  columns=['Bill.Hold Group', 'Claim.Jurisdiction State', 'Risk Status', 'Bill.Billing Type', 'Total Charges', 'Total Allowance'])

        # Convert categorical columns to string type
        categorical_columns = ['Bill.Hold Group', 'Claim.Jurisdiction State', 'Risk Status', 'Bill.Billing Type']
        for col in categorical_columns:
            input_data[col] = input_data[col].astype(str)

        encoded_columns = pd.DataFrame(encoder.transform(input_data[categorical_columns]).toarray()) # Apply the encoder to the categorical columns
        encoded_columns.columns = encoded_columns.columns.astype(str) # Convert the encoded column names to strings
        input_data = input_data.drop(categorical_columns, axis=1) # Drop the old categorical columns from original df
        input_data = input_data.join(encoded_columns) # Join the encoded df with the original df

        # Predict error probability on the new data
        input_pred = clf.predict(input_data)
        return input_pred

    def high_risk_notif(self):
        if self.checker == 1:
            # Clear bg
            self.lbl_errorpropensity.setStyleSheet("background-color: rgba(0,0,0,0); color: black; border: 1px; border-width: 1px; border-style: none; border-radius: 7px;")
            self.checker = 2
        elif self.checker == 2:
            # Red bg
            self.lbl_errorpropensity.setStyleSheet("background-color: red; color: white; border: 1px; border-width: 1px; border-style: solid; border-radius: 7px;")
            self.checker = 1

    def timer_event(self):
        if self.start_timer:
            self.time = self.time.addSecs(1)
        self.lbl_timer.setText(self.time.toString("hh:mm:ss"))

    def upload_allocation(self):
        # If ever we end up using something else to allocate the users' bills
        # Just make sure to use the same headers to avoid errors here
        # If not, then you might end up re-writing this function.

        self.checker = 0

        def open_file_dialog():
            file_dialog = QFileDialog(self)
            file_dialog.setWindowTitle("Select Allocation File...")
            file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
            file_dialog.setViewMode(QFileDialog.ViewMode.Detail)

            if file_dialog.exec():
                return file_dialog.selectedFiles()[0]
                
        xl_file = open_file_dialog()
        data_frame = pd.read_excel(xl_file, sheet_name="Main Bills Directory").to_json(orient="records")
        json_data = json.loads(data_frame)
        self.clear_all_fields()
        config = configparser.ConfigParser()
        config.read("config.ini")

        for item in json_data:
            try:
                user_nid = item["NAME"]
                user_nid = user_nid.split(",")[0].lower()

                if user_nid == self.username:
                    self.allocation.append(item)
            except:
                continue

        self.lbl_icn.setText(str(self.allocation[0]["CONTROL NUMBER"]))
        self.lbl_claimnumber.setText(self.allocation[0]["CLAIM.NUMBER"])
        self.lbl_holdgroup.setText(self.allocation[0]["BILL.HOLD GROUP"])
        self.lbl_jurisdiction.setText(self.allocation[0]["1. CLAIM.JURISDICTION STATE"])
        self.lbl_riskstatus.setText(self.allocation[0]["CURRENT RISK STATUS"])
        self.lbl_billingtype.setText(self.allocation[0]["BILL.BILLING TYPE"])
        self.lbl_totalcharges.setText("${:0,.2f}".format(self.allocation[0]["TOTAL CHARGES"]))
        self.lbl_totalallowance.setText("${:0,.2f}".format(self.allocation[0]["TOTAL ALLOWANCE"]))
        
        if config["BRAT_CONFIG"]["use_error_propensity"] == "true":
            print("Using Error Propensity model")
            # Make sure to update all the models first. For now, you can get the models from the workforce resource.
            # This takes up the following args, and should return "Error" or "No Error" based on that
            error_probability = self.error_propensity(self.allocation[0]["BILL.HOLD GROUP"],
                                                    self.allocation[0]["1. CLAIM.JURISDICTION STATE"], 
                                                    self.allocation[0]["CURRENT RISK STATUS"],
                                                    self.allocation[0]["BILL.BILLING TYPE"],
                                                    self.allocation[0]["TOTAL CHARGES"],
                                                    self.allocation[0]["TOTAL ALLOWANCE"])
            
            if error_probability[0] == "Error":
                self.checker = 2
                self.lbl_errorpropensity.setText("HIGH ERROR RISK")

                self.show_notification("!", "POTENTIAL ERROR RISK", "white", "blue")
            else:
                self.checker = 0
                self.lbl_errorpropensity.setText(" ")
        else:
            print("Using WFM Error Propensity")
            if self.allocation[0]["QA"] == "High":
                self.checker = 2
                self.lbl_errorpropensity.setText("HIGH ERROR RISK")

                self.show_notification("!", "POTENTIAL ERROR RISK", "white", "blue")
            else:
                self.checker = 0
                self.lbl_errorpropensity.setText(self.allocation[0]["QA"])

        # If you manage to think of somewhere to put these information in the UI. Just enable and assign to your labels/textboxes.
        # self.lbl_atriskdays.setText(str(self.allocation[0]["AT RISK DAYS (TARGET)"]))
        # self.lbl_criticaldays.setText(str(self.allocation[0]["CRITICAL DAYS (TARGET)"]))
        # self.lbl_projatriskdate.setText(pd.to_datetime(self.allocation[0]["PROJECTED AT RISK DATE"], unit="D", origin="1899-12-30").strftime("%m/%d/%Y"))
        # self.lbl_projcritdate.setText(pd.to_datetime(self.allocation[0]["PROJECTED CRITICAL RISK DATE"], unit="D", origin="1899-12-30").strftime("%m/%d/%Y"))
        # self.lbl_priority.setText(str(self.allocation[0]["PRIORITY"]))
        # self.lbl_inventorydate.setText(pd.to_datetime(self.allocation[0]["INVENTORY DATE"], unit="D", origin="1899-12-30").strftime("%m/%d/%Y"))
        # self.lbl_incomingleftover.setText(str(self.allocation[0]["INCOMING/LEFTOVER"]))
        QCoreApplication.processEvents()

        if not self.driver == None:
            self.extract_strataware()
        
    def next_allocation(self):
        current_index = 999

        for index, item in enumerate(self.allocation):
            if index >= len(self.allocation) - 1: break

            if str(item["CONTROL NUMBER"]) == self.lbl_icn.text():
                current_index = index
                break

        if current_index >= len(self.allocation):
            return
        
        self.clear_all_fields()
        current_index += 1
        config = configparser.ConfigParser()
        config.read("config.ini")

        self.lbl_icn.setText(str(self.allocation[current_index]["CONTROL NUMBER"]))
        self.lbl_claimnumber.setText(self.allocation[current_index]["CLAIM.NUMBER"])
        self.lbl_holdgroup.setText(self.allocation[current_index]["BILL.HOLD GROUP"])
        self.lbl_jurisdiction.setText(self.allocation[current_index]["1. CLAIM.JURISDICTION STATE"])
        self.lbl_riskstatus.setText(self.allocation[current_index]["CURRENT RISK STATUS"])
        self.lbl_billingtype.setText(self.allocation[current_index]["BILL.BILLING TYPE"])
        self.lbl_totalcharges.setText("${:0,.2f}".format(self.allocation[current_index]["TOTAL CHARGES"]))
        self.lbl_totalallowance.setText("${:0,.2f}".format(self.allocation[current_index]["TOTAL ALLOWANCE"]))
        
        if config["BRAT_CONFIG"]["use_error_propensity"] == "true":
            print("Using Error Propensity model")
            error_probability = self.error_propensity(self.allocation[0]["BILL.HOLD GROUP"],
                                                    self.allocation[0]["1. CLAIM.JURISDICTION STATE"],
                                                    self.allocation[0]["CURRENT RISK STATUS"],
                                                    self.allocation[0]["BILL.BILLING TYPE"],
                                                    self.allocation[0]["TOTAL CHARGES"],
                                                    self.allocation[0]["TOTAL ALLOWANCE"])
            
            if error_probability[0] == "Error":
                self.checker = 2
                self.lbl_errorpropensity.setText("HIGH ERROR RISK")

                self.show_notification("!", "POTENTIAL ERROR RISK", "white", "blue")
            else:
                self.checker = 0
                self.lbl_errorpropensity.setText(" ")
        else:
            print("Using WFM Error Propensity")
            if self.allocation[0]["QA"] == "High":
                self.checker = 2
                self.lbl_errorpropensity.setText("HIGH ERROR RISK")

                self.show_notification("!", "POTENTIAL ERROR RISK", "white", "blue")
            else:
                self.checker = 0
                self.lbl_errorpropensity.setText(self.allocation[0]["QA"])

        # self.lbl_atriskdays.setText(str(self.allocation[current_index]["AT RISK DAYS (TARGET)"]))
        # self.lbl_criticaldays.setText(str(self.allocation[current_index]["CRITICAL DAYS (TARGET)"]))
        # self.lbl_projatriskdate.setText(pd.to_datetime(self.allocation[current_index]["PROJECTED AT RISK DATE"], unit="D", origin="1899-12-30").strftime("%m/%d/%Y"))
        # self.lbl_projcritdate.setText(pd.to_datetime(self.allocation[current_index]["PROJECTED CRITICAL RISK DATE"], unit="D", origin="1899-12-30").strftime("%m/%d/%Y"))
        # self.lbl_priority.setText(str(self.allocation[current_index]["PRIORITY"]))
        # self.lbl_inventorydate.setText(pd.to_datetime(self.allocation[current_index]["INVENTORY DATE"], unit="D", origin="1899-12-30").strftime("%m/%d/%Y"))
        # self.lbl_incomingleftover.setText(str(self.allocation[current_index]["INCOMING/LEFTOVER"]))
        QCoreApplication.processEvents()

        if not self.driver == None:
            self.window_focus('SmartCoder')
            self.ui_to_front()
            self.extract_strataware()

    def bill_image(self):
        # The plan here is to open the bill image as you would the operative report.
        # Create a separate PDF viewer window and use the OCR text finder in ocr_tesseract.py
        # The end result is to highlight every information on Strataware that can be found in the bill image PDF.
        # Consult ops first on what should be included in that search.

        if len(self.driver.find_elements(By.XPATH, "//html/body[@id='SmartCoder']/label[@id='capBillInfo']/b/span/u")) == 0:
            ctypes.windll.user32.MessageBoxW(0, "Bill image not found", "Bill Review Assist", 0)
        else:
            self.driver.find_element(By.XPATH, "//html/body[@id='SmartCoder']/label[@id='capBillInfo']/b/span/u").click()

    def show_guidelines(self):
        # I do not recommend to include this function on future releases.
        # The file's too large, and have a bunch of comments on the cells so you might have a difficult time showing all the guidelines in the UI.
        # Just tell ops to manually open the file. And if you end up still using this, it should automatically navigate to the pricing state sheet
        # once they click on the Guidelines button. Remember, OPEN THE SHEET FIRST.

        try:
            state = "AL" if self.tbx_pricingstate.text() == "" else self.tbx_pricingstate.text()

            ExcelApp = win32.gencache.EnsureDispatch("Excel.Application")
            for workbook in ExcelApp.Workbooks:
                if "CRS_All_States" in workbook.Name:
                    wb = workbook
                    break

            self.ui_to_front()
            for ws in wb.Worksheets:
                if state.strip() in ws.Name and "2" in ws.Name:
                    ws.Activate()
                    break

            ExcelApp.Visible = True
            self.statusbar_update("Guidelines navigated to ")
        except:
            self.statusbar_update("Please open the CRS Guidelines first.")
            return

    def upload_cpt_aaos(self):
        # AAOS PDF uploader.
        # Just a simple copy-paste from a folder to the tool's folder.
        # If for some reason we encounter errors related to user privileges, just move the folder somewhere else.

        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("Select AAOS file...")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        file_dialog.setViewMode(QFileDialog.ViewMode.Detail)

        if file_dialog.exec():
            if os.path.exists(f"{os.getcwd()}\\resources\\aaos_latest.pdf"):
                os.remove(f"{os.getcwd()}\\resources\\aaos_latest.pdf")
            copy(file_dialog.selectedFiles()[0], f"{os.getcwd()}\\resources\\aaos_latest.pdf")

    def extract_aaos(self, cpt_code):
        fp = f"{os.getcwd()}\\resources\\aaos_latest.pdf"

        with open(fp, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            reader.decrypt("") # The AAOS file is encrypted without password

            omitted_txt1 = "Complete Global Service Data for Orthopaedic Surgery, 2019 edition, has been significantly \n"
            omitted_txt2 = "revised. Instructions for using the new global service data can be found in the introductory materials \n"
            omitted_txt3 = "on pages vii-xviii.\n"
            omitted_txt4 = "Complete Global Service Data for Orthopaedic Surgery copyright 2019 by the American Academy \n"
            omitted_txt5 = "of Orthopaedic Surgeons. All rights reserved.\n"
            omitted_txt6 = "Information contained in this document should not be construed as the only way in which contents \n"
            omitted_txt7 = "of a global service package can be determined. This document is explicitly not intended to present a standard of care or a practice parameter. Additional services may be requir"
            omitted_txt8 = "ed that are not listed."
            omitted_txt9 = "American Academy of Orthopaedic Surgeons\n"
            omitted_txt10 = "CPT codes, descriptions, and\n"
            omitted_txt11 = "numeric modifiers only are copyright\n"
            omitted_txt12 = "2018 American Medical Association.\n"
            omitted_txt13 = "All rights reserved."
            omitted_txt14 = "Complete Global Service Data for Orthopaedic Surgery"
            
            for pg in reader.pages:
                # Simple loop to search for the CPT code in the AAOS PDF
                if f"CPT Code: {cpt_code}" in pg.extract_text():
                    row_count = self.tbl_aaos.rowCount()
                    self.tbl_aaos.setRowCount(row_count + 1)

                    extracted_text = pg.extract_text()
                    output_text = extracted_text.replace(omitted_txt1, "").replace(omitted_txt2, "").replace(omitted_txt3, "") \
                        .replace(omitted_txt4, "").replace(omitted_txt5, "").replace(omitted_txt6, "").replace(omitted_txt7, "").replace(omitted_txt8, "") \
                        .replace(omitted_txt9, "").replace(omitted_txt10, "").replace(omitted_txt11, "").replace(omitted_txt12, "").replace(omitted_txt13, "") \
                        .replace(omitted_txt14, "").replace(f"CPT Code: {cpt_code}", "")
                    
                    self.tbl_aaos.setItem(row_count, 0, QTableWidgetItem(cpt_code))
                    self.tbl_aaos.setItem(row_count, 1, QTableWidgetItem(output_text))
                    self.tbl_aaos.verticalHeader().setSectionResizeMode(row_count, QHeaderView.ResizeToContents)
                    
                    break

    def show_info1(self):
        self.info1 = Info1Window(self.driver, self)
        self.info1.setWindowTitle("Bill Review Assist")
        self.info1.show()

        userid = os.getlogin()
        infos = self.read_text('config\\info1.txt')
        arr_infos = infos.split('\n')

        self.info1.comboBox_selection.addItem("")
        for info in arr_infos:
            info = info.replace('<ID>', userid)
            self.info1.comboBox_selection.addItem(info)

    def ocr_reader(self):
        def open_file_dialog():
            file_dialog = QFileDialog(self)
            file_dialog.setWindowTitle("Select PDF File...")
            file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
            file_dialog.setViewMode(QFileDialog.ViewMode.Detail)

            if file_dialog.exec():
                return file_dialog.selectedFiles()[0]

        self.op_report_pdf = open_file_dialog()
        self.op_report_pdf = self.op_report_pdf.replace("/", "\\")
        ocr_tesseract_update.create_initial_images(self.op_report_pdf)

        if self.imageviewer_window == None:
            self.imageviewer_window = ImageViewerWindow(self)

        self.imageviewer_window.img_index = 0
        self.imageviewer_window.load_image()
        self.imageviewer_window.show()

    def read_text(self, path):
        with open(path, 'r') as file:
            content = file.read()
        return content
        
if hasattr(Qt, 'AA_EnableHighDpiScaling'):
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

config = configparser.ConfigParser()
config.read("config.ini")

if "BRAT_CONFIG" in config:
    if config["BRAT_CONFIG"]["app_scaling_override"] == "2": # 0 -> Unchecked; 1 -> Partially Checked; 2 -> Checked
        os.environ["QT_SCALE_FACTOR"] = config["BRAT_CONFIG"]["app_scaling_override_value"]
    else:
        # Check screen scaling to correct app window scaling since the app window size is locked
        # If you can re-scale the widgets properly, you can ditch this
        app_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0)
        if app_scale <= 125:
            os.environ["QT_SCALE_FACTOR"] = "1.0"
            config["BRAT_CONFIG"]["app_scaling_override_value"] = "1"
        elif app_scale > 125:
            os.environ["QT_SCALE_FACTOR"] = "0.5"
            config["BRAT_CONFIG"]["app_scaling_override_value"] = "0.5"

        with open("config.ini", "w") as f:
            config.write(f)
else:
    config["BRAT_CONFIG"] = {
        "url": "https://strataware.lmig.com/",
        "app_scaling_override": "0",
        "app_scaling_override_value": "1",
        "use_error_propensity": "false",
        "gradient": "127", # Gradient to be used in OCR module
        "ocr_dpi": "225" # DPI to be used in generating PDF images
    }

    with open("config.ini", "w") as f:
        config.write(f)

if __name__ == "__main__":
    app = QApplication(argv)
    UIWindow = UI()
    app.exec()