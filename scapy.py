import xlrd
import xlsxwriter
import requests
from bs4 import BeautifulSoup
import io
from PyQt5.QtWidgets import QMessageBox

class scrapy_pic:
    def __init__(self, window, file_in, file_out):
        #self.file_in = file_in
        self.file_out = file_out + 'result.xlsx'
        self.window = window
        soursebook = xlrd.open_workbook(file_in)
        self.resultbook = xlsxwriter.Workbook(self.file_out)
        #all_sheet_names = soursebook.sheetnames()
        self.src_sheet =  soursebook.sheet_by_index(0)  #错误未处理
        self.result_sheet = self.resultbook.add_worksheet('sheet1')  #结果表格
        self.title_format = self.resultbook.add_format({'bold': True, 'align': 'center'})
        self.cell_format = self.resultbook.add_format({'align': 'center', 'valign': 'vcenter'})

    def preprocess(self):       #预处理数据
        self.result_sheet.set_column('A:B', 14)     #设置列宽
        self.result_sheet.set_column('C:D', 50)
        self.result_sheet.set_column('E:E', 20)
        self.result_sheet.set_column('F:G', 9)
        self.result_sheet.set_column('H:H', 19)
        row1 = self.src_sheet.row_values(0)
        self.result_sheet.set_row(0,17,self.title_format)
        self.result_sheet.write(0,0,row1[0])
        self.result_sheet.write(0,1,'宝贝图片')
        for td in range(1,len(row1)):
            self.result_sheet.write(0, td + 1, row1[td])


    def scrap(self):   #根据源表格的地址抓取对应图片
        if(self.src_sheet.nrows == 0):
            QMessageBox.information(self.window, 'Message', '表格为空')

        self.window.progress_bar.setRange(0, self.src_sheet.nrows)
        for rowind in range(1, self.src_sheet.nrows):  # src_sheet.nrows
            self.window.statusBar().showMessage('正在爬取图片' + str(rowind))
            tmp_row_content = []
            clonum = self.src_sheet.ncols
            for colind in range(0, self.src_sheet.ncols):
                tmp_row_content.append(self.src_sheet.cell_value(rowind, colind))

            response = requests.get(tmp_row_content[1])
            html = response.content.decode('gbk', errors='ignore')
            soup = BeautifulSoup(html, 'lxml')  # 初始化BeautifulSoup
            all_tbgallery = soup.find('div', class_="tb-gallery")  # 先找到最大的div
            pick = all_tbgallery.find('div', class_='tb-pic tb-s50')
            img = pick.find('img')
            picsrc = img['data-src']  # 此处提取的是50x50大小的图
            # result_picurl = picsrc[:-10]  #原图
            result_pic_url = picsrc.replace('50x50', '100x100')  # 把大小换成想要的大小
            if (result_pic_url.startswith('//')):
                result_pic_url = 'http:' + result_pic_url

            ima_data = requests.get(result_pic_url)
            data_stream = io.BytesIO(ima_data.content)
            self.result_sheet.set_row(rowind, 80, self.cell_format)
            self.result_sheet.write(rowind, 0, str(int(tmp_row_content[0])))
            self.result_sheet.insert_image(rowind, 1, result_pic_url, {'image_data': data_stream})
            self.result_sheet.write_row(rowind, 2, tmp_row_content[1:])
            self.window.progress_bar.setValue(rowind + 1)

        self.resultbook.close()
        QMessageBox.information(self.window, 'Message', '成功')
        self.window.statusBar().showMessage('结果表格生成' + self.file_out)
