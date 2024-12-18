from excel_operator import excel_operate
from userinterface import Application
if __name__ == "__main__":
    app = Application(excel_operate)
    app.mainloop()