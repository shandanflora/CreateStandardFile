from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

from mainwindow import *
from CreateStandardFile import *


class WorkThread(QThread):
    signal_info = pyqtSignal()
    signal_bar = pyqtSignal(int)

    def __init__(self, dict_para):
        super().__init__()
        self.dict = dict_para

    def run(self):
        # now_time = datetime.datetime.now()
        # print('2. %s' % now_time)
        # create_catalog.write_excel(self.dict, self.signal_bar)
        parseData = ParseData()
        parseData.readSrcFile(self.dict['src_file'])
        dict_cap = parseData.get_dict_cap()
        dict_res = parseData.get_dict_res()
        dict_other = parseData.get_dict_other()
        create_file = createStandardFile()
        create_file.write_excel(self.dict['lib_cap'],
                                self.dict['lib_res'],
                                self.dict['obj_file'],
                                dict_cap,
                                dict_res, dict_other)
        self.signal_info.emit()
        pass

# noinspection PyArgumentList


class main_window(QMainWindow):
    # constructor
    def __init__(self, name='Dlg'):
        super().__init__()
        self.label = QLabel()
        self.progressBar = QProgressBar()
        self.ui = Ui_MainWindow()
        self.initUI(name)

    # init function
    def initUI(self, name):
        self.ui.setupUi(self)
        self.statusBar().addPermanentWidget(self.label)
        self.statusBar().addPermanentWidget(self.progressBar)
        # This is simply to show the bar
        # self.progressBar.setGeometry(0, 0, 50, 15)
        self.progressBar.setRange(0, 100)  # 设置进度条的范围
        self.progressBar.setValue(0)
        self.setWindowTitle(name)
        self.initConnect()

    # initial slot and connect
    def initConnect(self):
        self.ui.btn_src.clicked.connect(lambda: self.btn_search_file_clicked(
                                        self.ui.lineEdit_src_file, "已选择源文件"))
        self.ui.btn_search_res.clicked.connect(lambda: self.btn_search_file_clicked(
                                        self.ui.lineEdit_res, "已选择电阻库"))
        self.ui.btn_search_cap.clicked.connect(lambda: self.btn_search_file_clicked(
                                        self.ui.lineEdit_cap, "已选择电容库"))
        self.ui.btn_obj_path.clicked.connect(self.path_btn_clicked)
        self.ui.btn_gen.clicked.connect(self.gen_btn_clicked)
        self.ui.btn_clear.clicked.connect(self.clear_btn_clicked)
        pass

    # ###########################################
    #        slot function
    # ###########################################
    def btn_search_file_clicked(self, lineEdit, string):
        file_name = QFileDialog.getOpenFileName(self.ui.centralwidget, 'open file', '/')
        lineEdit.setText(file_name[0])
        if file_name[0] != "":
            self.ui.statusbar.setStyleSheet("0;")
            self.ui.statusbar.setStyleSheet("font-size:15pt;")
            self.ui.statusbar.showMessage(string)

    def path_btn_clicked(self):
        file_name = QFileDialog.getExistingDirectory(self.ui.centralwidget, 'open folder', '/')
        self.ui.lineEdit_path.setText(file_name)
        # reset style sheet
        self.ui.statusbar.setStyleSheet("0;")
        self.ui.statusbar.setStyleSheet("font-size:15pt;")
        if file_name != "":
            file = self.ui.lineEdit_src_file[self.ui.lineEdit_src_file.rindex('/'):]
            file = file.index()
            self.ui.statusbar.showMessage("已选择目标路径")

    def gen_btn_clicked(self):
        if len(self.ui.lineEdit_src_file.text()) == 0:
            self.ui.statusbar.setStyleSheet("font-size:15pt;""background-color:#FF0000;")
            self.ui.statusbar.showMessage("源文件不能为空！！！")
        elif len(self.ui.lineEdit_res.text()) == 0:
            self.ui.statusbar.setStyleSheet("font-size:15pt;""background-color:#FF0000;")
            self.ui.statusbar.showMessage("电阻文件不能为空！！！")
        elif len(self.ui.lineEdit_cap.text()) == 0:
            self.ui.statusbar.setStyleSheet("font-size:15pt;""background-color:#FF0000;")
            self.ui.statusbar.showMessage("电容文件不能为空！！！")
        elif len(self.ui.lineEdit_path.text()) == 0:
            self.ui.statusbar.setStyleSheet("font-size:15pt;""background-color:#FF0000;")
            self.ui.statusbar.showMessage("目标路径不能为空！！！")
        elif len(self.ui.lineEdit_file_name.text()) == 0:
            self.ui.statusbar.setStyleSheet("font-size:15pt;""background-color:#FF0000;")
            self.ui.statusbar.showMessage("目标文件名不能为空！！！")
        else:

            obj_file = self.ui.lineEdit_path.text() + "/" + self.ui.lineEdit_file_name.text()
            src_file = self.ui.lineEdit_src_file.text()
            lib_cap = self.ui.lineEdit_cap.text()
            lib_res = self.ui.lineEdit_res.text()
            dict_para = {'src_file': src_file, 'obj_file': obj_file,
                         'lib_cap': lib_cap, 'lib_res': lib_res,
                         'bar': self.progressBar, 'label': self.label}
            self.label.setText("正在生成:")
            self.thread = WorkThread(dict_para)
            self.thread.signal_info.connect(self.update_info)
            self.thread.signal_bar.connect(self.update_bar)
            self.thread.start()

    def update_info(self):
        self.label.setText("")
        self.ui.statusbar.showMessage("已生成目录文件")
        self.progressBar.setValue(100)

    def update_bar(self, value):
        self.progressBar.setValue(value)
        pass

    def clear_btn_clicked(self):
        self.ui.statusbar.setStyleSheet("0;")
        self.ui.statusbar.setStyleSheet("font-size:15pt;")
        self.ui.statusbar.showMessage("清空编辑框")
        self.ui.lineEdit_src_file.setText("")
        self.ui.lineEdit_cap.setText("")
        self.ui.lineEdit_res.setText("")
        self.ui.lineEdit_path.setText("")
        self.ui.lineEdit_file_name.setText("")
        self.progressBar.setValue(0)
