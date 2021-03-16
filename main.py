from mainwindow_inherit import *

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = main_window("生成标准库")
    window.show()
    sys.exit(app.exec_())

