from email.message import EmailMessage
import os
import smtplib
from email.message import EmailMessage
import pandas as pd
import threading
import math

email_list = []
path ='C:/Users/USER/hello/qaquality'
try:
    os.chdir(path)
except FileNotFoundError:
    print("Directory: {0} does not exist".format(path))
except NotADirectoryError:
    print("{0} is not a directory".format(path))
except PermissionError:
    print("You do not have permissions to change to {0}".format(path))

email_list = []
passss = []
with open('emaillist.txt') as f:
    for line in f.readlines():
        if line.find('\n')>=0:
            line = line.replace("\n","")
        a = line.split(',')
        email_list.append(a[0])
        passss.append(a[1])



subject = 'อบรม Public หลักสูตร : CQI 9'#!ต้องเปลี่ยนนะ
# files=[]
files =['cqi/1.jpg','cqi/2.jpg','cqi/3.jpg']#!ต้องเปลี่ยนนะ
# hr_email = 'C:/Users/USER/hello/qaquality/testdata.xlsx'
# hr_email = 'C:/Users/USER/Downloads/qa_email_Nov_1.xlsx'#!ต้องเปลี่ยนนะ
hr_email = 'C:/Users/USER/Downloads/cqi.xlsx'#!ต้องเปลี่ยนนะ
# contentori = '''    
# PUBLIC  TRAINING 2023<br>
# <br>
# แผนอบรม Public Training ประจำปี 2023<br>
# <br>
# สนใจส่งพนักงานเข้าอบรมหลักสูตรไหน<br>
# <br>
# ติดต่อที่ คุณจิรภาQ&A<br>
# 027102138<br>
# 0858038380<br>
# qaquality@hotmail.com<br>
# '''
contentori = """
เรียน : สมาชิกครับ<br>
แจังข่าว  อบรม public<br>
==================<br>
หลักสูตร: CQI 9<br>
   Heat Treat  System<br>
   Assessment<br>
   4th  Edition<br>
<br>
วันที่: 3-05-2023<br>
เวลา: 9:00 -16:00<br>
สถานที: Novotel Hotel<br>
<br>
วิทยากร: อ.อเนก<br>
<br>
"ราคา2900บาทต่อท่าน"<br>
<br>
สนใจฝึกอบรม<br>
ติดต่อ คุณจิรภา Q&A<br>
027102138<br>
0858038380<br>
qaquality@hotmail.com<br>
<br>
"Note: รับจำนวนจำกัด"<br>
"""
#!ต้องเปลี่ยนนะ
def sendemail(df,sender_email,password,subject,files,i,content,sheetname,nono):
     with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, password)
        try:
            email = df.iloc[i, :]["email"]
            receiver_email = email.replace("\u200b", "")
            

            message = EmailMessage()
            message["From"] = sender_email
            message["To"] = receiver_email
            message["Subject"] = subject
            # message["Bcc"] = "qaquality@hotmail.com"


            message.add_alternative("""\
                <!DOCTYPE html>
                <html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office">
                <head>
                    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
                    <meta charset="utf-8">
                <meta name="viewport" content="width=device-width,initial-scale=1">
                <meta name="x-apple-disable-message-reformatting">
                <title></title>
                    <!--[if mso]>
                <style>
                    table {border-collapse:collapse;border-spacing:0;border:none;margin:0;}
                    div, td {padding:0;}
                    div {margin:0 !important;}
                    </style>
                <noscript>
                    <xml>
                    <o:OfficeDocumentSettings>
                        <o:PixelsPerInch>96</o:PixelsPerInch>
                    </o:OfficeDocumentSettings>
                    </xml>
                </noscript>
                <![endif]-->
                <style>
                    table, td, div, h1, p {
                    font-family: Arial, sans-serif;
                    }
                    @media screen and (max-width: 530px) {
                    .unsub {
                        display: block;
                        padding: 8px;
                        margin-top: 14px;
                        border-radius: 6px;
                        background-color: #555555;
                        text-decoration: none !important;
                        font-weight: bold;
                    }
                    .col-lge {
                        max-width: 100% !important;
                    }
                    }
                    @media screen and (min-width: 531px) {
                    .col-sml {
                        max-width: 27% !important;
                    }
                    .col-lge {
                        max-width: 73% !important;
                    }
                    }
                </style>
                </head>
                <body style="margin:0;padding:0;word-spacing:normal;background-color:#D0D0D0;">
                <div role="article" aria-roledescription="email" lang="en" style="text-size-adjust:100%;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;background-color:#D0D0D0;">
                    <table role="presentation" style="width:100%;border:none;border-spacing:0;">
                    <tr>
                        <td align="center" style="padding:0;">
                        <!--[if mso]>
                        <table role="presentation" align="center" style="width:600px;">
                        <tr>
                        <td>
                        <![endif]-->
                        <table role="presentation" style="width:94%;max-width:600px;border:none;border-spacing:0;text-align:left;font-family:Arial,sans-serif;font-size:16px;line-height:22px;color:#363636;">
                            <tr>
                            <td style="padding:0;font-size:24px;line-height:28px;font-weight:bold;">
                                <a href="http://www.qaquality.co.th/" style="text-decoration:none;"><img src="https://i.ibb.co/6mHg2tG/QA1-3.png" width="600" alt="" style="width:100%;height:auto;display:block;border:none;text-decoration:none;color:#363636;"></a>
                            </td>
                            </tr>
                            <tr>
                            <td style="padding:35px 30px 11px 30px;font-size:0;background-color:#ffffff;border-bottom:1px solid #f0f0f5;border-color:rgba(201,201,207,.35);">
                                <!--[if mso]>
                                <table role="presentation" width="100%">
                                <tr>
                                <td style="width:145px;" align="left" valign="top">
                                <![endif]-->
                                <div class="col-sml" style="display:inline-block;width:100%;max-width:145px;vertical-align:top;text-align:left;font-family:Arial,sans-serif;font-size:14px;color:#363636;">
                                <img src="https://i.ibb.co/W5VzHWx/aj.png" width="115" alt="" style="width:115px;max-width:80%;margin-bottom:20px;">
                                </div>
                                <!--[if mso]>
                                </td>
                                <td style="width:395px;padding-bottom:20px;" valign="top">
                                <![endif]-->
                                <div class="col-lge" style="display:inline-block;width:100%;max-width:395px;vertical-align:top;padding-bottom:20px;font-family:Arial,sans-serif;font-size:16px;line-height:22px;color:#363636;">
                                <p style="margin-top:0;margin-bottom:18px;">"""
                                +content+"""</p>
                                <p style="margin:0;"><a href="http://www.qaquality.co.th/contact.html" style="background: #ff3884; text-decoration: none; padding: 10px 25px; color: #ffffff; border-radius: 4px; display:inline-block; mso-padding-alt:0;text-underline-color:#ff3884"><!--[if mso]><i style="letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:20pt">&nbsp;</i><![endif]--><span style="mso-text-raise:10pt;font-weight:bold;">ติดต่อเรา</span><!--[if mso]><i style="letter-spacing: 25px;mso-font-width:-100%">&nbsp;</i><![endif]--></a></p>
                                </div>
                                <!--[if mso]>
                                </td>
                                </tr>
                                </table>
                                <![endif]-->
                            </td>
                            </tr>
                            <tr>
                            <td style="padding:10px;font-size:24px;line-height:28px;font-weight:bold;background-color:#ffffff;border-bottom:1px solid #f0f0f5;border-color:rgba(201,201,207,.35);">
                                <a style="text-decoration:none;"><img src="https://i.ibb.co/mR2t19j/QA1-4.png" width="600" alt="" style="width:100%;height:auto;border:none;text-decoration:none;color:#363636;"></a>
                            </td>
                            </tr>
                            <tr>
                            <td style="padding:30px;background-color:#ffffff;">
                                <p style="margin:0;">ขอขอบคุณทุกองค์กรที่ใช้บริการทั้งที่ปรึกษาและวิทยากรจากทีมงานของ Q&A</p>
                            </td>
                            </tr>
                            <tr>
                            <td style="padding:30px;text-align:center;font-size:12px;background-color:#404040;color:#cccccc;">
                                <p style="margin:0 0 8px 0;"><a href="https://www.facebook.com/QA-Quality-and-Calibration-Co-Ltd-176049912442479" style="text-decoration:none;"><img src="https://i.ibb.co/mSFKFn8/icon-facebook.png" width="40" height="40" alt="f" style="display:inline-block;color:#cccccc;"></a> <a href="https://line.me/ti/p/_slN0qEIeX" style="text-decoration:none;"><img src="https://i.ibb.co/pKzfqrs/line-icon.png" width="40" height="40" alt="t" style="display:inline-block;color:#cccccc;"></a></p>
                                <p style="margin:0;font-size:14px;line-height:20px;">&reg; Q&A Quality and Calibration,<br></p>
                                <p style="margin:0;font-size:14px;line-height:20px;">50/46,52 Moo 2, T. Bangkaew, A. Bangplee  Samutprakarn 10540 <br></p>
                                <p style="margin:0;font-size:14px;line-height:20px;">Tel. 02 710 2138, 02 710 2896<br></p>       
                                <p style="margin:0;font-size:14px;line-height:20px;">Email: qaquality@hotmail.com<br></p>
                            </td>
                            </tr>
                        </table>
                        <!--[if mso]>
                        </td>
                        </tr>
                        </table>
                        <![endif]-->
                        </td>
                    </tr>
                    </table>
                </div>
                </body>
                </html>
                """, subtype='html')
            # print(f"finish {i+1}")
            for file in files:
                with open(file,'rb') as f:
                    file_data = f.read()
                    file_name = f.name
                message.add_attachment(file_data, maintype='application',subtype ='octet-stream',filename=file_name)                
            
            
            server.send_message(message)
            print(f"sent {i+1} {sheetname} {nono}")

        except Exception as e: 
            print(e)
            pass
        server.quit()  


def sendingmarket(testdata,senderemail,passs,subject,files,sheetname,no=0):

    df = pd.read_excel(testdata,sheet_name=sheetname)
    # print(df) 
    sender_email = senderemail
    password = passs
    # print(df.head())
    if sheetname == "online Zoom" or "HR":
        totalsheet = math.ceil(df.shape[0]/450)
        # print(df.iloc[450*no:450*no+450])
        if sheetname == "online Zoom" and no >=46:
            return("out of index")
        if sheetname == "HR" and no >=13:
            return("out of index")
        df = df.iloc[450*no:450*no+450]
        # print(totalsheet)
        pass
    for i in range(df.shape[0]):
        # if i <35:
        #     continue
           
        sendemail(df,sender_email,password,subject,files,i,contentori,sheetname,no)
#!online Zoom 45 times+17 40
#!HR 12 times
#!Sheet7
th0 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[10],passss[10],subject,files,'online Zoom',3])
th1 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[11],passss[11],subject,files,'online Zoom',4])
th2 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[12],passss[12],subject,files,'online Zoom',5])
# th3 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[13],passss[13],subject,files,'online Zoom',3])
# th4 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[14],passss[14],subject,files,'online Zoom',4])
# th5 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[15],passss[15],subject,files,'online Zoom',5])

th0.start()
th1.start()
th2.start()
# th3.start()
# th4.start()
# th5.start()