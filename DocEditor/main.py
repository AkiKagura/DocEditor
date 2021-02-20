import sys
import win32com.client
from PyQt5.QtWidgets import *


class Example(QWidget):

    def __init__(self):
        super().__init__()

        self.hasOpen = 0
        self.initUI()

    def initUI(self):
        self.addr = QLineEdit()

        lbl1 = QLabel("Address:")
        lbl2 = QLabel("Function:")
        vbox = QVBoxLayout()
        hbox1 = QHBoxLayout()
        hbox2 = QHBoxLayout()
        self.cb = QComboBox()
        btn_in = QPushButton("Import")
        btn_out = QPushButton("Export")

        self.cb.addItems(["Test1"])
        self.cb.addItem("Original")
        # self.cb.currentIndexChanged.connect(self.cb_changed)
        hbox1.addWidget(lbl1)
        hbox1.addWidget(self.addr)
        hbox1.addWidget(btn_in)
        hbox2.addWidget(lbl2)
        hbox2.addWidget(self.cb)
        hbox2.addStretch(0)
        hbox2.addWidget(btn_out)

        btn_in.clicked.connect(self.button_import)
        btn_out.clicked.connect(self.button_export)

        vbox.addLayout(hbox1)
        vbox.addLayout(hbox2)

        self.setLayout(vbox)
        self.center()

        self.setWindowTitle('Docs Edit')
        self.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def button_import(self):
        fname, _ = QFileDialog.getOpenFileName(self, 'Open', 'E:\\TestDatas')
        self.addr.setText(fname)
        if fname != '':
            self.hasOpen = 1
            self.word = win32com.client.Dispatch('Word.Application')
            self.word.Visible = 0  # 后台运行
            self.word.DisplayAlerts = 0  # 不显示，不警告
            self.doc = self.word.Documents.Open(fname)  # 打开一个已有的word文档

    def button_export(self):
        if self.hasOpen == 1:
            if self.cb.currentText() == "Test1":
                myRange1 = self.doc.Range(0, 0)
                myRange1.InsertBefore('Hello word')

            self.doc.Save()  # 保存
            # doc.SaveAs('asdasd.doc')  # 另存为

            self.doc.Close()  # 关闭 word 文档
            # word.Documents.Close(wc.wdDoNotSaveChanges)  # 保存并关闭 word 文档
            self.word.Quit()  # 关闭 office
            print('OK')
            self.hasOpen = 0


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
