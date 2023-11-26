# This file is part of the program that calculates the shifts of the workers for a month, based on their preferences.
# The icon was generated from https://www.craiyon.com/image/WwLKHMVkTO6u2UxWCXprLw
#
# @Author: Pravitas Theocharis
# @Version: 1.0
#

from ortools.sat.python import cp_model
from PyQt6 import QtCore, QtGui, QtWidgets
import calendar
import os
import sys
import traceback
import xlsxwriter
import random


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(1082, 675)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayoutWidget = QtWidgets.QWidget(parent=self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(440, 10, 576, 60))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(3, 5, 5, 5)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        spacerItem = QtWidgets.QSpacerItem(
            248,
            17,
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Minimum,
        )
        self.horizontalLayout.addItem(spacerItem)
        self.label_3 = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout.addWidget(self.label_3)
        self.cb_year = QtWidgets.QComboBox(parent=self.verticalLayoutWidget)
        self.cb_year.setObjectName("cb_year")
        self.cb_year.addItem("")
        self.cb_year.addItem("")
        self.cb_year.addItem("")
        self.cb_year.addItem("")
        self.cb_year.addItem("")
        self.cb_year.addItem("")
        self.cb_year.addItem("")
        self.horizontalLayout.addWidget(self.cb_year)
        self.label = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.cb_month = QtWidgets.QComboBox(parent=self.verticalLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred, QtWidgets.QSizePolicy.Policy.Fixed
        )
        sizePolicy.setHorizontalStretch(50)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.cb_month.sizePolicy().hasHeightForWidth())
        self.cb_month.setSizePolicy(sizePolicy)
        self.cb_month.setObjectName("cb_month")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.cb_month.addItem("")
        self.horizontalLayout.addWidget(self.cb_month)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_6 = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_2.addWidget(self.label_6)
        self.extra_holidays = QtWidgets.QLineEdit(parent=self.verticalLayoutWidget)
        self.extra_holidays.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.extra_holidays.sizePolicy().hasHeightForWidth()
        )
        self.extra_holidays.setSizePolicy(sizePolicy)
        self.extra_holidays.setObjectName("extra_holidays")
        self.horizontalLayout_2.addWidget(self.extra_holidays)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.bt_calculate = QtWidgets.QPushButton(parent=self.centralwidget)
        self.bt_calculate.setGeometry(QtCore.QRect(570, 340, 81, 81))
        font = QtGui.QFont()
        font.setPointSize(32)
        self.bt_calculate.setFont(font)
        self.bt_calculate.setCursor(
            QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        )
        self.bt_calculate.setFocusPolicy(QtCore.Qt.FocusPolicy.ClickFocus)
        self.bt_calculate.setObjectName("bt_calculate")
        self.graphicsView_results = QtWidgets.QGraphicsView(parent=self.centralwidget)
        self.graphicsView_results.setEnabled(True)
        self.graphicsView_results.setGeometry(QtCore.QRect(670, 140, 391, 491))
        self.graphicsView_results.setObjectName("graphicsView_results")
        self.label_4 = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(30, 20, 391, 41))
        font = QtGui.QFont()
        font.setPointSize(20)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(20, 130, 601, 71))
        self.label_5.setObjectName("label_5")
        self.TextEdit_enter_data = QtWidgets.QPlainTextEdit(parent=self.centralwidget)
        self.TextEdit_enter_data.setGeometry(QtCore.QRect(20, 200, 531, 431))
        self.TextEdit_enter_data.setObjectName("TextEdit_enter_data")
        self.horizontalLayoutWidget_2 = QtWidgets.QWidget(parent=self.centralwidget)
        self.horizontalLayoutWidget_2.setGeometry(QtCore.QRect(19, 90, 1041, 41))
        self.horizontalLayoutWidget_2.setObjectName("horizontalLayoutWidget_2")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_2)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_2 = QtWidgets.QLabel(parent=self.horizontalLayoutWidget_2)
        self.label_2.setEnabled(False)
        font = QtGui.QFont()
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_3.addWidget(self.label_2)
        self.label_7 = QtWidgets.QLabel(parent=self.horizontalLayoutWidget_2)
        self.label_7.setEnabled(False)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Preferred
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_7.sizePolicy().hasHeightForWidth())
        self.label_7.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_3.addWidget(self.label_7)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(parent=MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1082, 21))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(parent=self.menubar)
        self.menu.setObjectName("menu")
        self.menuHelp = QtWidgets.QMenu(parent=self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(parent=MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.action = QtGui.QAction(parent=MainWindow)
        self.action.setObjectName("action")
        self.actionManual = QtGui.QAction(parent=MainWindow)
        self.actionManual.setObjectName("actionManual")
        self.menu.addAction(self.action)
        self.menuHelp.addAction(self.actionManual)
        self.menubar.addAction(self.menu.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        # Create the Exit action
        self.action.triggered.connect(sys.exit)

        # Create the Manual action (open a window that shows the details of the program in a view text box)
        self.actionManual.triggered.connect(self.show_manual)

        # Create the action for the calculate button
        self.bt_calculate.clicked.connect(self.on_click)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label_3.setText(_translate("MainWindow", "Έτος:"))
        self.cb_year.setItemText(0, _translate("MainWindow", "2023"))
        self.cb_year.setItemText(1, _translate("MainWindow", "2024"))
        self.cb_year.setItemText(2, _translate("MainWindow", "2025"))
        self.cb_year.setItemText(3, _translate("MainWindow", "2026"))
        self.cb_year.setItemText(4, _translate("MainWindow", "2027"))
        self.cb_year.setItemText(5, _translate("MainWindow", "2028"))
        self.cb_year.setItemText(6, _translate("MainWindow", "2029"))
        self.label.setText(_translate("MainWindow", "Μήνας:"))
        self.cb_month.setItemText(0, _translate("MainWindow", "Ιανουάριος"))
        self.cb_month.setItemText(1, _translate("MainWindow", "Φεβρουάριος"))
        self.cb_month.setItemText(2, _translate("MainWindow", "Μάρτιος"))
        self.cb_month.setItemText(3, _translate("MainWindow", "Απρίλιος"))
        self.cb_month.setItemText(4, _translate("MainWindow", "Μάιος"))
        self.cb_month.setItemText(5, _translate("MainWindow", "Ιούνιος"))
        self.cb_month.setItemText(6, _translate("MainWindow", "Ιούλιος"))
        self.cb_month.setItemText(7, _translate("MainWindow", "Αύγουστος"))
        self.cb_month.setItemText(8, _translate("MainWindow", "Σεπτέμβριος"))
        self.cb_month.setItemText(9, _translate("MainWindow", "Οκτώβριος"))
        self.cb_month.setItemText(10, _translate("MainWindow", "Νοέμβριος"))
        self.cb_month.setItemText(11, _translate("MainWindow", "Δεκέμβριος"))
        self.label_6.setText(
            _translate(
                "MainWindow",
                "Εάν ο μήνας έχει Πρόσθετες Αργίες παρακαλώ γράψτε χωρίζοντας με κόμμα (,)",
            )
        )
        self.bt_calculate.setText(_translate("MainWindow", "⏩"))
        self.label_4.setText(_translate("MainWindow", "ΥΠΟΛΟΓΙΣΜΟΣ ΥΠΗΡΕΣΙΩΝ"))
        self.label_5.setText(
            _translate(
                "MainWindow",
                "ΠΑΡΑΔΕΙΓΜΑ ΕΙΣΑΓΩΓΗΣ (Για περισσότερα: Help/Manual)\n"
                "Ονομα1-3-1-[1,4,5,6,7,13,15]\n"
                "Ονομα2-3-1-[]\n"
                "Ονομα3-3-1-[10,11,12]\n"
                "",
            )
        )
        self.label_2.setText(_translate("MainWindow", "ΕΙΣΑΓΩΓΗ ΔΕΔΟΜΕΝΩΝ"))
        self.label_7.setText(_translate("MainWindow", "ΠΡΟΕΣΚΟΠΙΣΗ ΑΠΟΤΕΛΕΣΜΑΤΩΝ"))
        self.menu.setTitle(_translate("MainWindow", "File"))
        self.menuHelp.setTitle(_translate("MainWindow", "Help"))
        self.action.setText(_translate("MainWindow", "Exit"))
        self.actionManual.setText(_translate("MainWindow", "Manual"))

    def show_manual(self):
        """Show the manual in a new window"""
        self.manual_window = QtWidgets.QMainWindow()
        self.manual_window.resize(800, 600)
        self.manual_window.setWindowTitle("Manual")
        self.manual_window.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        self.manual_window.setAttribute(QtCore.Qt.WidgetAttribute.WA_DeleteOnClose)
        self.manual_window.show()
        self.manual_window.activateWindow()
        self.manual_window.raise_()
        self.manual_window.activateWindow()
        self.manual_window.setFocus()

        self.manual_text = QtWidgets.QTextEdit(parent=self.manual_window)
        self.manual_text.setGeometry(QtCore.QRect(0, 0, 800, 600))
        self.manual_text.setReadOnly(True)
        self.manual_text.setText(
            """
            ===========================
            Οδηγίες Χρήσης Προγράμματος
            ===========================
            Το παρόν πρόγραμμα υπολογίζει τις υπηρεσίες των εργαζομένων για ένα μήνα.
            Περιορισμοί που υπολογίζονται:
            1. Κάθε εργαζόμενος κάνει υπηρεσίες εντός των ορίων που του έχουν οριστεί. Μεγιστο και ελάχιστο αριθμό υπηρεσιών.
            2. Κάθε μέρα έχει μια υπηρεσία.
            3. Κάθε εργαζόμενος δεν μπορεί να κάνει δύο μέρες συνεχόμενες.
            4. Κάθε εργαζόμενος κάνει από μια υπηρεσίες τις αργίες.
            5. Κάθε εργαζόμενος δεν κάνει υπηρεσία τις ημέρες που έχει ορίσει ως μη διαθέσιμες.
            6. Υπολογίζει ως αργία το Σάββατο και την Κυριακή. Ο χρήστης εισάγει άμα υπάρχουν πρόσθετες αργίες, όπως 28 Οκτωβρίου, 25 Μαρτίου κλπ.
            Το πρόγραμμα παρουσιάζει τα αποτελέσματα σε ένα αρχείο xlsx και σε ένα text box. Το αρχείο excel αποθηκεύεται στα Εγγραφα του Υπολογιστή.
            Τα αποτελέσματα που παρουσιάζονται είναι μια πιθανή λύση του προβλήματος και με την εκ νέου εκτέκλεση εμφανίζει μια διαφορετική λύση εφόσον υπάρχει.Εάν δεν υπάρχει λύση, τότε το πρόγραμμα εμφανίζει το αντίστοιχο μήνυμα.
            
            ===========================     
            Τα δεδομένα που απαιτούνται είναι τα εξής:
            1. Έτος
            2. Μήνας
            3. Πρόσθετες αργίες (εάν υπάρχουν)
                    >>>Σημείωση:οι αργίες χωρίζονται με κόμμα (,)<<<
                    ΠΑΡΑΔΕΙΓΜΑ: 21,30
            4. Εργαζόμενοι
                    >>>Απαιτείται η εισαγωγή των εργαζομένων στην παρακάτω μορφή:
            <ΟΝΟΜΑ>-<ΜΕΓΙΣΤΟΣ ΑΡΙΘΜΟΣ>-<ΕΛΑΧΙΣΤΟΣ ΑΡΙΘΜΟΣ>-[<ΚΕΝΗ_ΜΕΡΑ_1, ΚΕΝΗ_ΜΕΡΑ_2,...]>
            
            !!! ΠΡΟΣΟΧΗ !!!
            - ΔΕΝ ΠΡΕΠΕΙ ΝΑ ΥΠΑΡΧΕΙ ΚΑΠΟΙΟ ΚΕΝΟ [SPACE] ΑΝΑΜΕΣΑ ΣΤΑ ΔΕΔΟΜΕΝΑ ΓΙΑ ΚΑΘΕ ΕΡΓΑΖΟΜΕΝΟ
            - ΟΙ ΕΡΓΑΖΟΜΕΝΟΙ ΧΩΡΙΖΟΝΤΑΙ ΜΕ ENTER

            ΠΑΡΑΔΕΙΓΜΑ:
            =================
            ΠΕΡΙΚΛΗΣ-3-1-[1,4,5,6,7,13,15]
            ΑΡΙΣΤΕΙΔΗΣ-3-1-[]
            ΑΛΕΞΑΝΔΡΟΣ-3-1-[10,11,12]

            ***Σημείωση****: 
            *Οι κενές μέρες είναι προαιρετικές αλλά απαιτείτε η εισαγωγή των κενών αγκυλών. 
            Σε περίπτωση που συμπλήρωθούν κενές μέρες, ο εργαζόμενος δεν θα εργαστεί σε αυτές τις μέρες.

            *Οι κενές μέρες πρέπει να είναι αριθμοί από 1-31 και να χωρίζονται με κόμμα (,)

            *Εάν κάποιος εργοζόμενος πρέπει να δουλέψει ακριβώς ένα συγκεκριμένο αριθμό υπηρεσιών, τοτε
            αρκεί να συμπληρωθεί ο ίδιος αυτός αριθμός ως μέγιστος και ελάχιστος αριθμός υπηρεσιών.
            ΠΑΡΑΔΕΙΓΜΑ: ΠΕΤΡΟΣ-3-3-[]

            *Εάν κάποιος εργοζόμενος πρέπει να δουλέψει ακριβώς μια συγκεκριμένη ημέρα, τοτε
            αρκεί να συμπληρωθούν οι υπόλοιπες υμερομηνίες του μήνα ως μη διαθέσημες. Ο παρακάτω
            εργαζόμενος θα δουλέψει μόνο τις 1 και 10 του μήνα (έστω ότι ο εν λόγω μήνας έχει 30 μέρες).
            ΠΑΡΑΔΕΙΓΜΑ: ΧΑΡΗΣ-1-1-[2,3,4,5,6,7,8,9,11,12,13,14,15,16,17,20,21,22,23,24,25,26,27,28,29,30]


            ΔΕΔΟΜΕΝΑ ΓΙΑ ΔΟΚΙΜΕΣ:	
            =================
            ΟΝΟΜΑ_01-3-1-[1,4,5,6,7,13,15]
            ΟΝΟΜΑ_02-3-1-[]
            ΟΝΟΜΑ_03-3-1-[10,11,12]
            ΟΝΟΜΑ_04-3-1-[1,4,5,6,7,13,15]
            ΟΝΟΜΑ_05-3-1-[]
            ΟΝΟΜΑ_06-3-1-[20,21,22,23,24,25,26,27,28,29,30]
            ΟΝΟΜΑ_07-3-1-[1,4,5,6,7,13,15]
            ΟΝΟΜΑ_08-3-1-[]
            ΟΝΟΜΑ_09-3-1-[10,11,12]
            ΟΝΟΜΑ_10-3-1-[1,4,5,6,7,13,15]
            ΟΝΟΜΑ_11-4-1-[]
            ΟΝΟΜΑ_12-4-1-[10,11,12]
            """
        )
        self.manual_text.show()

    def get_month(self):
        month_dict = {
            "Ιανουάριος": 1,
            "Φεβρουάριος": 2,
            "Μάρτιος": 3,
            "Απρίλιος": 4,
            "Μάιος": 5,
            "Ιούνιος": 6,
            "Ιούλιος": 7,
            "Αύγουστος": 8,
            "Σεπτέμβριος": 9,
            "Οκτώβριος": 10,
            "Νοέμβριος": 11,
            "Δεκέμβριος": 12,
        }
        return month_dict[self.cb_month.currentText()]

    def get_year(self):
        return int(self.cb_year.currentText())

    def get_extra_holidays(self):
        extra_holidays = self.extra_holidays.text()
        if extra_holidays == "":
            return []
        extra_holidays_list = [int(holiday) for holiday in extra_holidays.split(",")]
        return extra_holidays_list

    def get_workers(self):
        text = self.TextEdit_enter_data.toPlainText()
        # chesk if there is any data
        if text == "":
            raise Exception("Δεν υπάρχουν δεδομένα")

        # check if there is any empty line or a line with spaces and delete it
        text = "\n".join([line.strip() for line in text.splitlines() if line != ""])
        text = "\n".join([line.strip() for line in text.splitlines() if line != " "])

        workers = []
        for id, line in enumerate(text.splitlines()):
            name, max, min, p = line.split("-")
            pref = [int(x) for x in p[1:-1].split(",") if x != ""] if p != "" else []
            workers.append(Worker(id, name, int(max), int(min), pref))
        return workers

    def update_ui(self, text):
        self.show_text(text)

    def on_click(self):
        try:
            # get calendar data
            month = self.get_month()
            year = self.get_year()
            try:
                extra_holidays = self.get_extra_holidays()
            except Exception as e:
                raise Exception("Λάθος στην εισαγωγή των επιπλέον αργιών")

            # get workers data
            workers = self.get_workers()

            # create schedule
            schedule = Schedule(year, month, extra_holidays)

            # delete previous results
            schedule.delete_result()

            # add workers to schedule
            for worker in workers:
                schedule.add_worker(worker)

            # calculate schedule
            results = schedule.calculate_schedule()
            filename = schedule.export_to_excel()

            if results is None:
                raise Exception("Δεν βρέθηκε λύση")
            else:
                output_text = "\n".join(
                    [
                        f"\nΤο αναλυτικό xlsx αρχείο αποθηκεύτηκε στον φάκελο:\n {filename}"
                    ]
                    + [""]
                    + [schedule.return_results(results)]
                )
                self.update_ui(output_text)

        except Exception as e:
            self.show_text("Σφάλμα: " + str(e))
            traceback.print_exc()

    def show_text(self, text):
        scene = QtWidgets.QGraphicsScene()
        text_item = scene.addText(text)
        text_item.setPos(0, 0)
        self.graphicsView_results.setScene(scene)


class Worker:
    def __init__(self, id, name, max=3, min=1, pref=[]):
        self.id = id
        self.name = name
        self.max = max
        self.min = min
        self.pref = pref
        self.total = 0
        self.total_holidays = 0
        self.shifts = []

    def delete(self):
        if hasattr(self, "id"):
            del self.id
        if hasattr(self, "name"):
            del self.name
        if hasattr(self, "max"):
            del self.max
        if hasattr(self, "min"):
            del self.min
        if hasattr(self, "pref"):
            del self.pref
        if hasattr(self, "total"):
            del self.total
        if hasattr(self, "total_holidays"):
            del self.total_holidays
        if hasattr(self, "shifts"):
            del self.shifts

    def __repr__(self):
        return f"{self.name} ({self.min}-{self.max})"


class Schedule:
    def __init__(self, year, month, custom_days=[]):
        self.year = year
        self.month = month
        self.workers = []
        self.custom_days = custom_days
        self.shifts = {}

    def add_worker(self, worker):
        self.workers.append(worker)

    def delete_result(self):
        if hasattr(self, "shifts"):
            del self.shifts
        for worker in self.workers:
            worker.delete()

    def find_holidays(self):
        # Get the calendar for the specified month and year
        cal = calendar.monthcalendar(self.year, self.month)

        # Define the day numbers for Sunday and Saturday
        sunday = calendar.SUNDAY
        saturday = calendar.SATURDAY

        # Use list comprehension to generate the list of days
        days = [
            day for week in cal for day in (week[sunday], week[saturday]) if day != 0
        ]

        # Add custom days to the list if provided
        if self.custom_days:
            days.extend(self.custom_days)

        return sorted(set(days))

    def calculate_schedule(self):
        holidays = self.find_holidays()
        # active_shifts_per_day = 1
        num_days = calendar.monthrange(self.year, self.month)[1] + 1

        # Create the model
        model = cp_model.CpModel()

        # randomize the results
        random.shuffle(self.workers)

        # Create the variables
        shifts = {
            (worker.name, day): model.NewBoolVar(f"{worker.name}_{day}")
            for worker in self.workers
            for day in range(1, num_days)
        }

        # Create the constraints
        for worker in self.workers:
            # Each worker must work a maximum of x shifts
            model.Add(
                sum(shifts[(worker.name, day)] for day in range(1, num_days))
                <= worker.max
            )

            # Each worker must work a minimum of x shifts
            model.Add(
                sum(shifts[(worker.name, day)] for day in range(1, num_days))
                >= worker.min
            )

            # Each worker must not work more than 1 shift in a row
            for day in range(1, num_days - 1):
                model.Add(
                    shifts[(worker.name, day)] + shifts[(worker.name, day + 1)] <= 1
                )

            # Each worker must work at most  one time per month on holidays
            model.Add(
                sum(
                    shifts[(worker.name, day)]
                    for day in holidays
                    if day not in worker.pref
                )
                <= 1
            )

        # Each day must have exactly 1 worker
        for day in range(1, num_days):
            model.AddExactlyOne([shifts[(worker.name, day)] for worker in self.workers])

        # Each worker must not work on their preferences
        model.Maximize(
            sum(
                shifts[(worker.name, day)]
                for worker in self.workers
                for day in range(1, num_days)
                if day not in worker.pref
            )
        )

        # Create the solver
        solver = cp_model.CpSolver()

        # Solve the model
        status = solver.Solve(model)

        # Check if the model is feasible
        if status == cp_model.FEASIBLE or status == cp_model.OPTIMAL:
            # Create the results
            results = {}
            for day in range(1, num_days):
                for worker in self.workers:
                    if solver.Value(shifts[(worker.name, day)]) == 1:
                        results[day] = worker.name
                        worker.shifts.append(day)
                        break

            # Count the shifts
            for worker in self.workers:
                worker.total = sum(
                    solver.Value(shifts[(worker.name, day)])
                    for day in range(1, num_days)
                )
                worker.total_holidays = sum(
                    solver.Value(shifts[(worker.name, day)]) for day in holidays
                )

            self.shifts = results
            return results
        else:
            return None

    def return_results(self, results):
        # Create a list of formatted strings for each worker's information
        worker_info = [
            f"{worker.name} >> {worker.total} / {worker.max} and {worker.total_holidays} on holidays"
            for worker in sorted(self.workers, key=lambda x: x.id)
        ]

        # Sort the results and format them
        result_lines = [
            f"{day:02d}: {worker}" for day, worker in sorted(results.items())
        ]

        # total holidays
        total_holidays = sum(worker.total_holidays for worker in self.workers)

        # Join the result lines and worker information
        output_text = "\n".join(
            worker_info
            + [""]
            + result_lines
            + [""]
            + [f"Total holidays: {str(total_holidays)}"]
        )
        # return the optimized text
        return output_text

    def export_to_excel(self):
        holidays = self.find_holidays()
        num_days = (
            calendar.monthrange(self.year, self.month)[1] + 1
        )  # +1 because the range is 0-30 and we need 1-31

        # save the file to the my documents folder
        filename = os.path.expanduser(
            f"~\Documents\schedule_{self.year}_{self.month}.xlsx"
        )
        # if the file already exists, add a number to the end of the filename
        if os.path.exists(filename):
            i = 1
            while os.path.exists(filename):
                filename = os.path.expanduser(
                    f"~\Documents\schedule_{self.year}_{self.month} ({i}).xlsx"
                )
                i += 1

        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()

        # Define the formats for the cells
        format_title = workbook.add_format({"bold": True})
        format_pref_day = workbook.add_format(
            {"bold": True, "bg_color": "#ffcccc", "border": 1}
        )
        format_hodiday = workbook.add_format(
            {"bold": True, "bg_color": "#d3d3d3", "border": 1}
        )
        format_shift_day = workbook.add_format({"border": 1, "align": "center"})

        # Set the column width
        worksheet.set_column("A:A", 20)  # Name
        worksheet.set_column("B:B", 10)  # Total Shifts
        worksheet.set_column("C:C", 10)  # Total Holidays
        worksheet.set_column("D:AI", 3)  # days of the month

        # Write the headers
        worksheet.write("A2", "ΟΝΟΜΑ", format_title)
        worksheet.write("B2", "ΣΥΝΟΛΟ", format_title)
        worksheet.write("C2", "ΑΡΓΙΕΣ", format_title)

        # sort the workers by id
        self.workers = sorted(self.workers, key=lambda x: x.id)

        # set the row and column to start writing the data
        row = 2
        col = 0

        # Write employee names and excel formulas for total shifts and total holidays
        for worker in self.workers:
            worksheet.write(row, col, str(worker.name))
            worksheet.write(row, col + 1, f'COUNTIF(E{row+1}:AI{row+1};"X")')
            worksheet.write(
                row, col + 2, f'COUNTIFS(E{row+1}:AI{row+1};"X";E$1:AI$1;TRUE)'
            )
            row += 1

        col = 4

        # Write the days as headers
        for day in range(num_days - 1):
            day_label = day + 1
            if day_label in holidays:
                header_format = format_hodiday
                worksheet.write(0, col + day, "TRUE", format_hodiday)
            else:
                header_format = format_title
            worksheet.write(1, col + day, day_label, header_format)

        # check if every day has a shift
        worksheet.write(
            len(self.workers) + 3,
            col + day,
            f'COUNTIF(E3:AI{len(self.workers)+3};"X")',
        )

        row = 2
        row_pref = 2

        # Write the schedule for each worker
        for worker in self.workers:
            worker_schedule = [
                "X" if day in worker.shifts else " " for day in range(1, num_days)
            ]
            print(worker_schedule)
            worksheet.write_row(row, col, worker_schedule, format_shift_day)
            row += 1

            # make the cells with the prefrenced days light red
            for day in worker.pref:
                worksheet.write(row_pref, col + day - 1, "", format_pref_day)
            row_pref += 1

        # Close the workbook.
        workbook.close()

        return filename


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()

    sys.exit(app.exec())
