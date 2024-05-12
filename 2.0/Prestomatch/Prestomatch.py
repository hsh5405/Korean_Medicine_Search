import sys
import re
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QFont, QColor
from PyQt5.QtWidgets import QMainWindow, QApplication, QDesktopWidget, QMessageBox, QVBoxLayout, QCheckBox, QDialog, QLabel, QLineEdit, QSplitter, QPushButton, QPlainTextEdit, QTableWidget, QTableWidgetItem, QHeaderView, QGridLayout, QWidget, QStatusBar

#https://www.youtube.com/watch?v=Zd7EkceyoUM = pip install auto-py-to-exe = just find and run
#pyinstaller --onefile --noconsole --icon=logo.ico --add-data "logo.ico;." --add-data "standard.xlsx;." --add-data "prescription.xlsx;." Prestomatch.py
#pyinstaller Prestomatch.spec

class Prestomatch(QMainWindow):
    def __init__(self):
        super().__init__()
        self.data_standard = {}
        self.data_prescriptions = {}
        self.part_prescriptions = {}
        self.max_rows = 0
        self.max_columns = 0
        self.icon_filename = "logo.ico"
        self.png_filename = "./logo.png"
        self.standard_filename = "standard.xlsx"
        self.prescription_filename = "prescription.xlsx"
        self.log_state = False
        
        self.initUI()
        self.load_standard_file()
        self.load_prescription_file()
    
    def log(self, text):
        if self.log_state:
            self.logTextEdit.appendPlainText(text)

    def load_standard_file(self, filename=None):
        try:
            if filename == None:
                filename = self.standard_filename
            
            data = {}
            
            sheets = list(pd.ExcelFile(filename).sheet_names)
            self.log(f"[{sys._getframe().f_code.co_name}]: Sheets of '{filename}' - {str(sheets)}.")
            
            df = pd.read_excel(filename, sheet_name=sheets[0], dtype='unicode')
            result = df.fillna("").values
            
            for i in result:
                if "H" in i[4]:
                    single_prescription = {
                        "type": i[2],
                        "disease": i[3]
                    }
                    data[i[4]] = single_prescription
            
            self.data_standard = data
            self.log(f"[{sys._getframe().f_code.co_name}]: '{self.standard_filename}' loaded.")
            
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")

    def load_prescription_file(self, filename=None):
        try:
            if filename == None:
                filename = self.prescription_filename
                
            data = {}
            
            sheets = list(pd.ExcelFile(filename).sheet_names)
            self.log(f"[{sys._getframe().f_code.co_name}]: Sheets of '{filename}' - {str(sheets)}.")
            
            df = pd.read_excel(filename, sheet_name=sheets[0], dtype='unicode')
            result = df.fillna("").values

            current_id = 0
            for i in result:
                if i[0] != '':
                    current_id, name = i[0].strip().split(".") 
                    single_prescription = {
                        "name": name,
                        "codes": sorted(i[1].replace("\n", " ").split()),
                        "ingredients": [i[3].strip()]#[[i[2].strip(), i[3].strip()]] only name not code
                    }
                    data[current_id] = single_prescription
                else:
                    if not "_" in i[2]:
                        data[current_id]["ingredients"].append(i[3].strip())#append([i[2].strip(), i[3].strip()])
            
            self.data_prescriptions = data
            self.log(f"[{sys._getframe().f_code.co_name}]: '{self.prescription_filename}' loaded.")
            
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")
    
    #standard.xlsx - find disease codes from prescription code 
    def find_standard(self, codes):
        try:
            def namecode(name):
                if name == "월경통": return "월"
                elif name == "안면신경마비": return "안"
                elif name == "뇌혈관질환 후유증": return "뇌"
                elif name == "알레르기성비염": return "비"
                elif name == "요추추간판탈출증": return "요"
                elif name == "기능성소화불량": return "소"
                else: return "?"
            
            types = []
            disease = []
            for code in codes:
                for key, value in self.data_standard.items():
                    if key == code:
                        types.append(value["type"])
                        disease.append(namecode(value["disease"]))
                        
            return f"({str(disease)})"
        
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")
    
    #prescription.xlsx - find all prescription codes with searching ingredient list from input_text   
    def find_prescription(self, input_texts):
        try:
            for key, value in self.data_prescriptions.items():
                ingredient_list = value.get("ingredients", [])
                self.max_columns = max(self.max_columns, len(ingredient_list))
                correct_states = [False for i in range(len(input_texts))]
                
                for i in range(len(input_texts)):
                    for ingredient in ingredient_list:
                        if input_texts[i] in ingredient:
                            correct_states[i] = True
                            
                if all(correct_states):
                    self.part_prescriptions[key] = value

            self.max_rows = len(self.data_prescriptions)
            self.log(f"[{sys._getframe().f_code.co_name}]: Data found.")
            
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")
    
    def initUI(self):
        try:
            self.setWindowTitle('Prestomatch 2.0')
            self.setWindowIcon(QIcon(self.png_filename))
            self.setGeometry(0, 0, 1600, 800)
            self.center_window()

            self.alert_window()
            self.create_widgets()
            self.create_layout()
            
            self.log(f"[{sys._getframe().f_code.co_name}]: UI loaded.")
            
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")

    def center_window(self):
        try:
            qr = self.frameGeometry()
            cp = QDesktopWidget().availableGeometry().center()
            qr.moveCenter(cp)
            self.move(qr.topLeft())
            
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")

    def alert_window(self):
        try:            
            dialog = QDialog(self)
            dialog.setWindowTitle('Copyright')
            dialog.setWindowIcon(QIcon(self.png_filename))

            layout = QVBoxLayout(dialog)

            alert_text = QLabel("ⓒ 1999 대한한의학회소문학회: All right reserved. \n무료배포용: 한의원내처방용입니다. 유료배포 및 기타 상업적이용을 금지합니다. \n\n"+
                                "<제작자 정보>\nGithub: https://github.com/hsh5405 \nE-mail: hsh5405@unist.ac.kr \n")
            layout.addWidget(alert_text)

            switch = QCheckBox('로그 기능 활성화')
            switch.setChecked(False)  # Initial state
            layout.addWidget(switch)

            ok_button = QPushButton('확인')
            ok_button.clicked.connect(dialog.accept)
            layout.addWidget(ok_button)
            
            dialog.exec_()
            
            self.log_state = switch.isChecked()
        
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")

    def create_widgets(self):
        try:
            font = QFont("맑은 고딕", 16)
            font.setBold(True)

            self.label = QLabel("검색어 입력: ")
            self.label.setFont(font)

            self.lineEdit = QLineEdit("")
            self.lineEdit.setFont(font)
            self.lineEdit.returnPressed.connect(self.button_search)
            
            self.btn_search = QPushButton("검색")
            self.btn_search.clicked.connect(self.button_search)
            self.btn_search.setFont(font)
            
            self.btn_info = QPushButton(f"도움말")
            self.btn_info.clicked.connect(self.button_info)
            self.btn_info.setFont(font)
            
            self.tableWidget = QTableWidget()

            self.statusBar = QStatusBar()
            self.statusBar.setFont(font)
            
            if self.log_state:
                self.logTextEdit = QPlainTextEdit()
                self.logTextEdit.setReadOnly(True)
        
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")

    def create_layout(self):
        try:
            widgets = QWidget(self)
            layout = QGridLayout(widgets)
            layout.addWidget(self.label, 0, 0)
            layout.addWidget(self.lineEdit, 0, 1)
            layout.addWidget(self.btn_search, 0, 2)
            layout.addWidget(self.btn_info, 0, 3)
            layout.addWidget(self.tableWidget, 1, 0, 1, 4)
            layout.addWidget(self.statusBar, 2, 0, 1, 4)
            
            if self.log_state:
                widgets2 = QWidget(self)
                layout2 = QGridLayout(widgets2)
                layout2.addWidget(self.logTextEdit)

            #split widgets through spliter
            splitter = QSplitter(Qt.Horizontal)
            splitter.addWidget(widgets)
            if self.log_state:
                splitter.addWidget(widgets2)

            #widgets size and ratio
            splitter.setSizes([800, 400])  
            splitter.setStretchFactor(0, 1)  
            splitter.setStretchFactor(1, 0)  

            self.setCentralWidget(splitter)
            
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")

    def button_search(self):
        try:
            self.part_prescriptions = {}
            self.max_rows = 0
            self.max_columns = 0
            self.tableWidget.clearContents()
            
            self.log(f"[{sys._getframe().f_code.co_name}]: Searching [{self.lineEdit.text()}].")
            
            if self.data_prescriptions == {}:
                self.statusBar.showMessage('엑셀 파일 데이터 2개를 프로그램과 같은 폴더에 배치해주세요.')
                return 0
            
            #Searching text encoding
            input_text = re.sub(r'\s+', ' ', self.lineEdit.text())
            input_texts = input_text.split(' ')
            
            if input_texts == ['']:
                self.statusBar.showMessage('검색어를 입력해주세요.')
                return 0
            
            self.find_prescription(input_texts)
            
            if self.part_prescriptions != {}:
                #Set bottom log text
                self.statusBar.showMessage('현 검색어: '+self.lineEdit.text())
                
                #Set table row and column
                self.tableWidget.setRowCount(self.max_rows)
                self.tableWidget.setColumnCount(self.max_columns+1) #first name column added 1 

                #Set table first row title
                tag = ["처방명"]+[f"한약재{i}" for i in range(self.max_columns)]
                self.tableWidget.setHorizontalHeaderLabels(tag)
                
                font = QFont("맑은 고딕", 12)
                font.setBold(True)
                color = QColor("#FFFF00")
                
                for i, item in enumerate(self.part_prescriptions.items()):
                    key, value = item
                    current_prescription = value["ingredients"]
                    
                    status = self.find_standard(value["codes"])
                    
                    self.tableWidget.setItem(i, 0, QTableWidgetItem(f"{key}.{value['name']}{status}"))
                    
                    for j in range(len(current_prescription)):
                        self.tableWidget.setItem(i, j+1, QTableWidgetItem(current_prescription[j]))
                        for input_text in input_texts:
                            if input_text in current_prescription[j]:
                                self.tableWidget.item(i, j+1).setFont(font)
                    #In last, set bold and color into first column
                    self.tableWidget.item(i, 0).setFont(font)
                    self.tableWidget.item(i, 0).setBackground(color)
                
                self.tableWidget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
                self.tableWidget.show()
            else:
                self.statusBar.showMessage("검색어를 입력하지 않았거나 올바르지 않은 검색어가 있습니다.")
        
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")
    
    def button_info(self):
        try:
            alert = QMessageBox()
            alert.setWindowTitle("ⓘ 도움말")
            alert.setWindowIcon(QIcon(self.png_filename))
            alert.setText(f"프로그램 사용 시, 제공되는 2개 파일({self.standard_filename}, {self.prescription_filename})을 같은 이름과 확장자로 실행기와 같은 폴더 내에 두어야 합니다.\n\n"+
                          "이 프로그램은 각 엑셀 파일의 1번째 시트만 사용되어 제작되었습니다.\n\n"+
                          "모두 기존 파일의 이름과 내부 형식을 준수한다면, 추가하거나 삭제하는 등의 수정이 가능합니다.\n")
            alert.addButton("확인", QMessageBox.AcceptRole)
            alert.setFixedSize(600, 200)
            alert.exec_()
        
        except Exception as e:
            print(e)
            self.log(f"[{sys._getframe().f_code.co_name}]: \n {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Prestomatch()
    ex.show()
    sys.exit(app.exec_())
    