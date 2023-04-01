import ssl
import shutil, os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from PIL import ImageFont,ImageDraw,Image
import time
def sending(testdata,senderemail,passs,pic,date,topic,folder):

    t1 = time.perf_counter()
    df = pd.read_excel(testdata)#ต้องเปลี่ยนนะ
    context = ssl.create_default_context()
    sender_email = senderemail
    password = passs
    folder = folder #ต้องเปลี่ยนนะ
    
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        # print("hi")
        for i, x in enumerate(df["num"]):
            # if i < 3:
            #     continue
            try:
                name = (df.iloc[i, :]["name"]).replace("\u200b", "")
                no = (df.iloc[i, :]["cert_num"])
                email = df.iloc[i, :]["email"]
                image = Image.open(pic) #ต้องเปลี่ยนนะ    
                draw = ImageDraw.Draw(image)


                
                date_email =date #ต้องเปลี่ยนนะ
                topic= topic #ต้องเปลี่ยนนะ
                
                certi = "Q "+str(no)+"/2023"
                

                font = ImageFont.truetype("TIMESBD0.TTF", 100)
                font2 = ImageFont.truetype("TIMESBD0.TTF", 80)
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
                
                
                receiver_email = email.replace("\u200b", "")
                


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
                
                
                
                
                server.sendmail(sender_email, receiver_email, text)
                print("sent {}".format(no))
                shutil.move(filename, folder)
            except KeyboardInterrupt:
                break
            except:
                pass
        server.quit()
    print(time.perf_counter() -t1)