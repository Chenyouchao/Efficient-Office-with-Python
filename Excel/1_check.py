'''
    疫情期间, 一些学校需要每天统计学生信息, 
    此脚本可自动找出未提交信息的学生. 
    
    1. 设置excel文件路径和名称, 并将学生信息按students_all字典排列;
    2. 输入python check.py 运行程序;
    3. 返回未提交信息的学生及其联系方式. 
'''

import pandas as pd
import time

# 设置excel文件目录及文件名
file_path = "D:\download\\"
file_name = "check.xls"
# 获取当前时间
time_ = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))

students_all = {"张抗疫": "18866668888", 
                "李爱国": "16688886666",
                }

try:
    df = pd.read_excel(file_path + file_name)
    students_done = df.loc[:, "姓名"]

    num = 0     # 未完成填写的人数
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
            print("%s,未完成数据填写！电话：%s-%s-%s" % 
            	(student, students_all[student][0:3], 
                students_all[student][3:7], 
                students_all[student][7:11]))

    if num == 0:
        print("\n全部学生完成填写, 时间：" + time_)
    else:
        print("\n共有%d个学生未完成填写！时间：" % num + time_)
except:
    print("读取文件失败···")



'''
    补充: for循环的一般写法 (效率较低)

    for i in students_list.keys():
        flag = False
        for j in name_list:
            if i == j:
                flag = True
        if flag is False:
            num += 1
            print("%s,未完成数据填写！电话：%s" % (i, students_list[i]))
'''