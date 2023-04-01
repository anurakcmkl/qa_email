import qadataextract
import qanew1
import os
import threading

date = '22/11/2022'#!ต้องเปลี่ยนนะ
topic = "การจัดทำบริบทขององค์กรและการกำหนดความต้องการ/ความคาดหวังผู้มีส่วนได้ส่วนเสีย"#!ต้องเปลี่ยนนะ
folder = '2155'#!ต้องเปลี่ยนนะ
pic = 'original2023(11).jpg'#!ต้องเปลี่ยนนะ
# url = "C:/Users/USER/Downloads/ข้อกำหนดISO500012018(การตอบกลับ).xlsx"
url = 'https://docs.google.com/spreadsheets/d/18OLJXboPxxx3BC99MZ60THD2FNa-GzUiQy6F1rskNj8/edit?usp=sharing'#!ต้องเปลี่ยนนะ
round =27#!ต้องเปลี่ยนนะ
folder_data = 'iso999_22_11_2022'#!ต้องเปลี่ยนนะ
codee = 'Q&A01'#!ต้องเปลี่ยนนะ
certlastnum= 30313#!ต้องเปลี่ยนนะ

path ='C:/Users/USER/hello/qaquality'
try:
    os.chdir(path)
except FileNotFoundError:
    print("Directory: {0} does not exist".format(path))
except NotADirectoryError:
    print("{0} is not a directory".format(path))
except PermissionError:
    print("You do not have permissions to change to {0}".format(path))


if not os.path.isdir(folder):
        os.mkdir(folder)
if not os.path.isdir(folder_data):
        os.mkdir(folder_data)


email_list = []
passss = []
with open('emaillist.txt') as f:
    for line in f.readlines():
        if line.find('\n')>=0:
            line = line.replace("\n","")
        a = line.split(',')
        email_list.append(a[0])
        passss.append(a[1])
# passss= 'qa1234567890'
# print(email_list)
# print(passss)

# qadataextract.qadata(url,round,folder_data,codee,450,certlastnum)

# th0 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_1.xlsx',email_list[0],passss[0],pic,date,topic,folder])
# th1 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_2.xlsx',email_list[1],passss[1],pic,date,topic,folder])
# th2 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_3.xlsx',email_list[2],passss[2],pic,date,topic,folder])
# th3 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_4.xlsx',email_list[3],passss[3],pic,date,topic,folder])
# th4 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_5.xlsx',email_list[4],passss[4],pic,date,topic,folder])
# th5 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_6.xlsx',email_list[5],passss[5],pic,date,topic,folder])
# th6 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_7.xlsx',email_list[6],passss[6],pic,date,topic,folder])
# th7 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_8.xlsx',email_list[7],passss[7],pic,date,topic,folder])
# th8 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_9.xlsx',email_list[14],passss[14],pic,date,topic,folder])

# th0.start()
# th1.start()
# th2.start()
# th3.start()
# th4.start()
# th5.start()
# th6.start()
# th7.start()
# th8.start()

