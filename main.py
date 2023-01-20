from pdf2docx import Converter
from PyQt6 import QtWidgets
import sys

class widget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('pdf2word')
        self.resize(350, 200)
        self.ui()

    def ui(self):
        self.txt_content = QtWidgets.QTextEdit(self)
        self.txt_content.setGeometry(20,20,300,120)

        self.btn_file = QtWidgets.QPushButton(self)
        self.btn_file.setText('选择文件')
        self.btn_file.move(20, 150)
        self.btn_file.clicked.connect(self.open_files) # 调用open_files方法

        self.btn_convert = QtWidgets.QPushButton(self)
        self.btn_convert.setText('开始转换')
        self.btn_convert.move(100, 150)
        self.btn_convert.clicked.connect(self.convert) # 调用convert方法

    def open_files(self):
        # filePath 为文件完整路径
        global file_path
        file_path = QtWidgets.QFileDialog.getOpenFileNames(self, '选择文件','', "PDF Files (*.pdf)")[0] # 可选多个文件

    def convert(self):
        # 单个文件
        if len(file_path) == 1 and file_path[0].split('.')[1] == 'pdf': 
            filename = pdf2word(file_path[0])
            print('文件个数：%s' % len(file_path))
            print('转换成功')
            print('文件保存位置：%s' % filename)
            self.txt_content.append('转换成功 %s' % filename)
        # 多个文件
        elif len(file_path) >= 2:
            print('文件个数：%s' % len(file_path))
            for file in file_path:
                filename = pdf2word(file)
                print('转换成功')
                print('文件保存位置：', filename)
                self.txt_content.append('转换成功 %s' % filename)

def pdf2word(filename):
    doc_file = filename.split('.')[0] + '.docx' # 获取原始文件名
    p2w = Converter(filename)
    p2w.convert(doc_file)
    p2w.close()
    return doc_file

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = widget() # 实例化部件
    window.show() # 展示控件
    sys.exit(app.exec()) # 在主线程中退出

if __name__ == '__main__': # 主入口
    main()