import qadataextract
import qanew1
import os
import threading

date = '23/03/2023'#!ต้องเปลี่ยนนะ
topic = "ISO45001 2018"#!ต้องเปลี่ยนนะ
folder = '9854'#!ต้องเปลี่ยนนะ
pic = 'original2023(03).jpg'#!ต้องเปลี่ยนนะ
# url = "C:/Users/USER/Downloads/ข้อกำหนด ISO 14001 2023-01-24 (การตอบกลับ).xlsx"
url = 'https://docs.google.com/spreadsheets/d/1kuhDDbReoBSMfZtOwGdIu4s4sAjOyMVceyB5YctVtuI/edit?usp=sharing'#!ต้องเปลี่ยนนะ
round =31#!ต้องเปลี่ยนนะ
folder_data = 'ISO45001_23_03_2023'#!ต้องเปลี่ยนนะ
codee = 'Q&A45'#!ต้องเปลี่ยนนะ
certlastnum= 7175#!ต้องเปลี่ยนนะ

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
th4 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_5.xlsx',email_list[4],passss[4],pic,date,topic,folder])
th5 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_6.xlsx',email_list[5],passss[5],pic,date,topic,folder])
th6 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_7.xlsx',email_list[6],passss[6],pic,date,topic,folder])
# th7 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_8.xlsx',email_list[7],passss[7],pic,date,topic,folder])

# th0 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_9.xlsx',email_list[8],passss[8],pic,date,topic,folder])
# th1 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_10.xlsx',email_list[9],passss[9],pic,date,topic,folder])

# th2 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_11.xlsx',email_list[10],passss[10],pic,date,topic,folder])
# th3 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_12.xlsx',email_list[11],passss[11],pic,date,topic,folder])
# th4 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_13.xlsx',email_list[12],passss[12],pic,date,topic,folder])
# th5 = threading.Thread(target=qanew1.sending,args=[f'{folder_data}/testdata{round}_14.xlsx',email_list[13],passss[13],pic,date,topic,folder])

# th0.start()
# th1.start()
# th2.start()
# th3.start()
th4.start()
th5.start()
th6.start()
# th7.start()

