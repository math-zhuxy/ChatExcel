from excel_operator import excel_operate
from userinterface import Application
from PyQt5.QtWidgets import QApplication
import sys
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Application(excel_operate)
    window.show()
    sys.exit(app.exec_())