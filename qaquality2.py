import ssl
import errorche
import shutil, os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from PIL import ImageFont,ImageDraw,Image
df = pd.read_excel("testdata18_2.xlsx")#ต้องเปลี่ยนนะ
for i, x in enumerate(df["num"]):
    # if i < 4:
    #     continue
    try:
        name = (df.iloc[i, :]["name"]).replace("\u200b", "")
        no = (df.iloc[i, :]["cert_num"])
        email = df.iloc[i, :]["email"]
        image = Image.open("original33.jpg") #ต้องเปลี่ยนนะ    
        draw = ImageDraw.Draw(image)


        
        date_email ="27/01/2022" #ต้องเปลี่ยนนะ
        topic= "ISO 9001:2015" #ต้องเปลี่ยนนะ
        folder = "096" #ต้องเปลี่ยนนะ
        certi = "Q "+str(no)+"/2022"
        

        font = ImageFont.truetype("TIMESBD0.TTF", 100)
        font2 = ImageFont.truetype("TIMESBD0.TTF", 80)
        draw.text((970, 2550), certi, font=font2, fill="black")
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
        sender_email = "qaquality111@gmail.com"
        sender_email2 = "qaquality145@gmail.com"
        receiver_email = email.replace("\u200b", "")
        password = "qa1234567890"


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

        if i%2 == 0:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, text)
            print("qa 1 sent {}".format(no))
            shutil.move(filename, folder)
        else:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email2, password)
                server.sendmail(sender_email2, receiver_email, text)
            print("qa 2 sent {}".format(no))
            shutil.move(filename, folder)
    except:
        pass
        errorche.errorcheck(i,name,email,no)
errorche.create(errorche.error_all)