import sys, time
import pandas as pd
from PySide2.QtWidgets import (QApplication, QWidget, QLabel, 
QPushButton, QTextEdit, QGridLayout, QVBoxLayout, QFileDialog)

class Check(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        grid = QGridLayout()

        self.lbl = QLabel()
        # self.lbl.resize(100, 10)
        self.btn = QPushButton('选择excel文件')
        self.btn.clicked.connect(self.showDialog)
        grid.addWidget(self.lbl, 1, 1, 1, 3)      # 选中(1, 1)到(1, 3)三个格子
        grid.addWidget(self.btn, 1, 4)            # 选中格子(1, 4)

        vbox = QVBoxLayout()

        self.text = QTextEdit()
        vbox.addLayout(grid)        # 区分: addLayout, addWidget
        vbox.addWidget(self.text)

        self.setLayout(vbox)
        self.setGeometry(300, 300, 600, 400)
        self.setWindowTitle("检查未填表学生")
        self.show()

    def showDialog(self):
        
        fname = QFileDialog.getOpenFileName(self, '选择excel文件', 'D:\下载', 'Excel file(*.xls)')

        if fname[0]:

            stus, time_ = self.checkStudents(fname[0])
            self.text.setText(stus)
            self.lbl.setText(time_)

    def checkStudents(self, fname):

        file_path = fname

        # 获取当前时间
        time_ = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))

        students_all = {"张抗疫": "18866668888", 
                        "李爱国": "16688886666",
                        }

        try:
            df = pd.read_excel(file_path)
            students_done = df.loc[:, "姓名"]

            num = 0     # 未完成填写的人数
            stus = ""   # 未完成填写的名单
            for student in students_all.keys():
                # for-else语句, 中途退出不执行else, 完整执行循环才进入else
                for student_done in students_done:
                    # 若在students_done列表中找到此student
                    # 则break退出内层循环, 查找下一个student
                    if student == student_done:
                        break
                else:
                    num += 1
                    # 输出电话号码格式为: xxx-xxxx-xxxx
                    stus += ("%s,未完成数据填写！电话：%s-%s-%s\n" % 
                        (student, students_all[student][0:3], 
                        students_all[student][3:7], 
                        students_all[student][7:11]))

            if num == 0:
                stus = "\n全部学生完成填写, 时间：" + time_
            else:
                stus += "\n共有%d个学生未完成填写！时间：" % num + time_
        except:
            stus = "读取文件失败···"

        return stus, time_


if __name__ == "__main__":

    app = QApplication(sys.argv)
    win = Check()
    app.exec_()