from tkinter import *
from tkinter import filedialog
import ssl
import os
# import errorche
import shutil, os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from PIL import ImageFont,ImageDraw,Image,ImageTk

root = Tk()
root.title("Qaquality Email System")
# root.iconbitmap('qalogo.ico')
root.geometry("800x400")
def browse_excel():
    global Excel_location
    root.filename = filedialog.askopenfilename(initialdir="/",title="Choose your Excel file",filetypes=(("Excel file","*.xlsx"),("Any file", "*.*")))
    Excel_location = root.filename
    excel_label = Label(root, text=Excel_location).place(x=250,y=20)

def browse_pic():
    global pic_location
    root.filename = filedialog.askopenfilename(initialdir="/",title="Choose your Picture file",filetypes=(("jpg file","*.jpg"),("Any file", "*.*")))
    pic_location = root.filename
    pic_label = Label(root, text=pic_location).place(x=250,y=60)
    if not len(pic_location)<2:
        my_pic =ImageTk.PhotoImage(Image.open(pic_location).resize((150,230),Image.ANTIALIAS))
        my_pic_label=Label(image=my_pic)
        my_pic_label.place(x=600,y=10)
        my_pic_label.image=my_pic

def browse_font():
    global font_location
    root.filename = filedialog.askopenfilename(initialdir="/",title="Choose your font file",filetypes=(("TTF file","*.TTF"),("Any file", "*.*")))
    font_location = root.filename
    font_label = Label(root, text=font_location).place(x=250,y=100)

def start():
    global folder, date, topic, email, password
    B_excel['state'] = 'disabled'
    B_font['state'] = 'disabled'
    B_Picture['state'] = 'disabled'
    Start_button['state'] = 'disabled'
    folder_entry['state'] = 'disabled'
    date_entry['state'] = 'disabled'
    topic_entry['state'] = 'disabled'
    email_entry['state'] = 'disabled'
    pass_entry['state'] = 'disabled'
    folder = folder_entry.get()
    date = date_entry.get()
    topic = topic_entry.get()
    emaill = email_entry.get()
    password = pass_entry.get()
    folder_label = Label(root, text=folder).place(x=400,y=140)
    date_label = Label(root, text=date).place(x=400,y=180)
    topic_label = Label(root, text=topic).place(x=400,y=220)
    email_label = Label(root, text=emaill).place(x=400,y=260)
    # password_label = Label(root, text=password).place(x=400,y=300)
    f= open("emailandpass.txt","w+")
    f.write(emaill+','+password)
    f.close
    if not os.path.isdir(folder):
        os.mkdir(folder)
    df = pd.read_excel(Excel_location)
    for i, x in enumerate(df["num"]):
        # if i > 0:
        #     continue
        try:
            name = (df.iloc[i, :]["name"]).replace("\u200b", "")
            no = (df.iloc[i, :]["cert_num"])
            email = df.iloc[i, :]["email"]
            image = Image.open(pic_location) #ต้องเปลี่ยนนะ    
            draw = ImageDraw.Draw(image)


            
            date_email =date #ต้องเปลี่ยนนะ
            topic= topic #ต้องเปลี่ยนนะ
            folder = folder #ต้องเปลี่ยนนะ
            certi = "Q "+str(no)+"/2023"
            

            font = ImageFont.truetype(font_location, 100)
            font2 = ImageFont.truetype(font_location, 80)
            draw.text((970, 2580), certi, font=font2, fill="black")
            draw.text((150, 1160), name, font=font, fill="#1017b4")
            image.save("cert_{}.jpg".format(no))
            print("Saved {}".format(no))
            subject = "Certificate "+topic+" "+date_email+" cert"+str(no)
            body = """เรียนผู้เกี่ยวข้อง
        ทาง Q&Aส่ง Certificate การฝึกอบรม """+topic+""" วันที่ """+date_email+"\n\n\n"+"""Q&A Quality and Calibration Co.,Ltd.
    50/46,52 Moo 2, T.Bangkaew, A.Bang Phli, Samut Prakarn, Thailand 10540
    Tel. 02 7102138, 02 7102896
    Fax. 02 7537452
    Mobile  085 8038380"""
            sender_email = emaill
            receiver_email = email.replace("\u200b", "")
            password = password


            message = MIMEMultipart()
            message["From"] = sender_email
            message["To"] = receiver_email
            message["Subject"] = subject
            message["Bcc"] = "qaquality@hotmail.com"


            message.attach(MIMEText(body, "plain"))

            filename = "cert_{}.jpg".format(no)


            with open(filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            
            encoders.encode_base64(part)

            
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {filename}",
            )

            
            message.attach(part)
            text = message.as_string()
            
            
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, text)
            print("qa 1 sent {}".format(no))
            shutil.move(filename, folder)
           
        except:
            pass
            
    


ExcelLabel = Label(root, text="เลือกไฟล์Excel")
ExcelLabel.place(x=30,y=20)
B_excel = Button(root, text ="Browse", command = browse_excel)
B_excel.place(x=200,y=20)

PictureLabel = Label(root, text="เลือกไฟล์รูป")
PictureLabel.place(x=30,y=60)
B_Picture = Button(root, text ="Browse", command = browse_pic)
B_Picture.place(x=200,y=60)

FontLabel = Label(root, text="เลือกไฟล์Font")
FontLabel.place(x=30,y=100)
B_font = Button(root, text ="Browse", command = browse_font)
B_font.place(x=200,y=100)

folder_entry = Entry(root,width=30)
folderLabel = Label(root, text="ใส่ชื่อแฟ้ม")
folderLabel.place(x=30,y=140)
folder_entry.place(x=200,y=140)

date_entry = Entry(root,width=30)
dateLabel = Label(root, text="ใส่วันที่")
dateLabel.place(x=30,y=180)
date_entry.place(x=200,y=180)

topic_entry = Entry(root,width=30)
topicLabel = Label(root, text="ใส่หัวข้อเรื่อง")
topicLabel.place(x=30,y=220)
topic_entry.place(x=200,y=220)

email_entry = Entry(root,width=30)
emailLabel = Label(root, text="ใส่email")
emailLabel.place(x=30,y=260)
email_entry.place(x=200,y=260)

pass_entry = Entry(root,width=30,show='*')
passLabel = Label(root, text="ใส่password")
passLabel.place(x=30,y=300)
pass_entry.place(x=200,y=300)

if os.path.isfile('emailandpass.txt'):
    with open('emailandpass.txt', 'r') as f:
        tempdata = f.read().split(',')
        email_temp = tempdata[0]
        pass_temp = tempdata[1]
        email_entry.insert(0, email_temp)
        pass_entry.insert(0, pass_temp)
        # print(email_temp)
        # print(pass_temp)


Start_button = Button(root,text="Start",command=start,width=20)
Start_button.place(x=360,y=350)

root.mainloop()
