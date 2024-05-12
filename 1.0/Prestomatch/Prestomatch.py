import sys
import pandas
from PyQt5.QtGui import QIcon, QFont, QColor
from PyQt5.QtWidgets import QMainWindow, QApplication, QDesktopWidget, QMessageBox, QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QGridLayout, QWidget, QStatusBar

df = pandas.read_excel("standard_data.xls", sheet_name='Sheet1', dtype='unicode')
data1 = df.fillna("").values

df = pandas.read_excel("standard_data.xls", sheet_name='Sheet2', dtype='unicode')
data2 = df.fillna("").values

data2 = [[0, i[0], i[1], i[2]] for i in data2]

#공통리스트 채우기
recent = ['', '']
for i in range(len(data2)):
    if data2[i][1] != '':
        recent[0] = data2[i][1]
    if data2[i][2] != '':
        recent[1] = data2[i][2]
    data2[i][0] = i
    data2[i][1] = recent[0]
    data2[i][2] = recent[1]

row_list = [i[3] for i in data2]

#처방별 공통질환 추가
for i in range(len(data2)):
    if data2[i][2] == '질환고유':
        index = list(filter(lambda x: row_list[x] == data2[i][3], range(len(row_list))))
        for j in index:
            data2[i].append(data2[j][1])
print(data2)


def searcher(word): #단어 단어 단어 ---
    word = word.split(" ")
    word_len = len(word)
    target_list = []
    for row in data1:
        boolean = 0
        for t in word:
            if t in row:
                boolean += 1
        if boolean == word_len:
            target_list.append(row)

    return [target_list, word]

def kind(word):
    if word in row_list:
        id = row_list.index(word)
        a = ''
        tag = data2[id][2]
        if tag == '질환고유':
            for i in range(4, len(data2[id])):
                if i == 4:
                    a += data2[id][i]
                else:
                    a += ", "+data2[id][i]
        else:
            a = tag
        word = word+"("+a+")"
    return word

class Prestomatch(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self): 
        self.setWindowTitle('Prestomatch')
        self.setWindowIcon(QIcon('logo.png'))
        self.setGeometry(0, 0, 1200, 800)
        #center
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

        alert = QMessageBox()
        alert.setWindowTitle('Copyright')
        alert.setWindowIcon(QIcon('logo.png'))
        alert.setText('ⓒ 1999 대한한의학회소문학회: All right reserved. \n무료배포용: 한의원내처방용입니다. 유료배포 및 기타 상업적이용을 금지합니다. \n\nGithub: \nhttps://github.com/leaonblue')
        alert.exec_()

        self.label = QLabel("입력: ")
        font = QFont("맑은 고딕", 16)
        font.setBold(True)
        self.label.setFont(font)

        self.lineEdit = QLineEdit("")
        self.lineEdit.setFont(font)

        self.btn = QPushButton("검색")
        self.btn.clicked.connect(self.btn_clicked)
        self.btn.setFont(font)
        
        self.tableWidget = QTableWidget()

        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        font = QFont("맑은 고딕", 9)
        font.setBold(True)
        self.statusBar.setFont(font)
        
        widgets = QWidget(self)
        layout = QGridLayout(widgets)
        layout.addWidget(self.label, 0, 0)
        layout.addWidget(self.lineEdit, 0, 1)
        layout.addWidget(self.btn, 0, 2)
        layout.addWidget(self.tableWidget, 1, 0, 1, 3)
        layout.addWidget(self.statusBar, 2, 0, 1, 3)
        self.setCentralWidget(widgets)

    def btn_clicked(self):
        text = searcher(self.lineEdit.text())
        result = text[0]
        word = text[1]

        if result != []:
            self.statusBar.showMessage('현 검색어: '+self.lineEdit.text())

            self.tableWidget.setRowCount(len(result))
            self.tableWidget.setColumnCount(len(result[0]))

            tag = ["처방명"]
            for i in range(len(result[0])-1):
                tag.append("한약재")
            self.tableWidget.setHorizontalHeaderLabels(tag)
            
            for i in range(len(result)):
                for j in range(len(result[0])):
                    input_text = str(result[i][j])
                    if j == 0:
                        input_text = kind(input_text)
                    self.tableWidget.setItem(i, j, QTableWidgetItem(input_text))
                    if input_text in word:
                        font = QFont("맑은 고딕", 10)
                        font.setBold(True)
                        self.tableWidget.item(i, j).setFont(font)
                font = QFont("맑은 고딕", 10)
                font.setBold(True)
                self.tableWidget.item(i, 0).setFont(font)
                self.tableWidget.item(i, 0).setBackground(QColor("#FFFF00"))
        else:
            self.statusBar.showMessage("검색어를 입력하지 않았거나 올바르지 않은 검색어가 있습니다.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Prestomatch()
    ex.show()
    sys.exit(app.exec_())