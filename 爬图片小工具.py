import sys
import scapy
from PyQt5.QtWidgets import QWidget,QMainWindow, QApplication,QPushButton,QMessageBox,\
    QLineEdit,QFileDialog,QGridLayout,QProgressBar

class Example(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        src_btn = QPushButton("选择文件", self)
        #src_btn.setGeometry(100, 70, 120, 60)
        result_btn = QPushButton("生成位置", self)
        self.ed1 = QLineEdit(self)
        self.ed2 = QLineEdit(self)

        gridLayoutWidget = QWidget(self)
        gridLayoutWidget.setGeometry(30, 30, 340, 180)
        gridlayout = QGridLayout(gridLayoutWidget)
        gridlayout.addWidget(src_btn,0 , 0, 1, 1)
        gridlayout.addWidget(self.ed1, 0, 1, 1, 1)
        gridlayout.addWidget(result_btn,1 , 0, 1, 1)
        gridlayout.addWidget(self.ed2, 1, 1, 1, 1)
        #self.setLayout(gridlayout)

        gene_btn = QPushButton("开始", self)
        gene_btn.setGeometry(50, 200, 80, 40)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(140, 210, 230, 22)


        src_btn.clicked.connect(self.choosefile)
        result_btn.clicked.connect(self.choose_dir)
        gene_btn.clicked.connect(self.generate_result)


        self.setGeometry(300, 300, 400, 300)
        self.setWindowTitle('爬取图片')
        self.show()


    def choosefile(self):
        fname = QFileDialog.getOpenFileName(self, '选择文件','', 'Excel Files (*.xls;*.xlsx)')
        self.progress_bar.reset()
        if not fname[0]:
            QMessageBox.information(self, 'warning', '请选择正确的源文件')

        self.ed1.setText(fname[0])
        ind = fname[0].rfind('/')
        folder = fname[0][:ind]
        self.ed2.setText(folder)
        self.statusBar().showMessage('源文件已选择')

    def choose_dir(self):
        dir_choose = QFileDialog.getExistingDirectory(self, "选取文件夹", self.ed2.text())  #选择保存路径
        self.ed2.setText(dir_choose)

    def generate_result(self):
        filein = self.ed1.text()
        file_out = self.ed2.text()
        if(not file_out.endswith('/')):
            file_out += '/'

        scra = scapy.scrapy_pic(self, filein, file_out)
        scra.preprocess()
        scra.scrap()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())