from PyQt5.QtWidgets import *
import pandas as pd
import sys
import mysql.connector as mysql
class Window(QWidget):
  def Convert_to_Database(self):
      read=pd.read_excel(self.excel_file.text(), engine='openpyxl')
      read_array=read.to_numpy()
      
      try:
        con=mysql.connect(user=self.username.text(),password=self.password.text(),host=self.hostname.text())
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setText("Loged into the database successfully")
        msgBox.setWindowTitle("Authentification")
        msgBox.setStandardButtons( QMessageBox.Cancel)
        msgBox.exec()
        for reading in read_array:
           finalvalues=[]
           finalvalues.append(str(reading[0]))
           finalvalues.append(str(reading[1]))
           finalvalues.append(str(reading[2]))
           finalvalues.append(str(reading[3]))
           finalvalues.append(str(reading[4]))
           finalvalues.append(str(reading[5]))
           finalvalues.append(str(reading[6]))
           cursor=con.cursor()
           cursor.execute("USE db4")
           sql=" INSERT INTO db4(serialnumber,entrynumber,volumenumber,district,year,user,hospital) VALUES(%s,%s,%s,%s,%s,%s,%s)"
           print(finalvalues)
           cursor.execute(sql,finalvalues)
           con.commit()
   
      except mysql.Error as e:
            msgBox = QMessageBox()
            msgBox.setIcon(QMessageBox.Information)
            msgBox.setText(str(e))
            msgBox.setWindowTitle("Authentification")
            msgBox.setStandardButtons( QMessageBox.Cancel)
            msgBox.exec()
          
  def __init__(self):
      super().__init__()
      self.setWindowTitle("Excel Converter")
      self.resize(700,700)
      self.setStyleSheet("background-color: green;color:red;")
      self.setMaximumSize(700,700)
      
      layout=QVBoxLayout()
      
      self.username=QLineEdit()
      self.username.setPlaceholderText("User Name")
      self.username.setStyleSheet("background-color: yellow")
      self.hostname=QLineEdit()
      self.hostname.setPlaceholderText("Host Name")
      self.hostname.setStyleSheet("background-color: yellow")
      self.password=QLineEdit()
      self.password.setPlaceholderText("Password ")
      self.password.setStyleSheet("background-color: yellow")
      self.excel_file=QLineEdit()
      self.excel_file.setStyleSheet("background-color: yellow")
      self.excel_file.setPlaceholderText("Full Path to Excel File")
      self.convert=QPushButton("Convert Excel To Database")
      self.convert.setStyleSheet("background-color:blue")
      self.convert.resize(40,40)
      layout.addWidget(self.username)
      layout.addWidget(self.hostname)
      layout.addWidget(self.password)
      layout.addWidget(self.excel_file)
      layout.addWidget(self.convert)
      self.setLayout(layout)
      self.convert.clicked.connect(self.Convert_to_Database)
      
  

          
       



     
       




if __name__=="__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    app.exec_()


