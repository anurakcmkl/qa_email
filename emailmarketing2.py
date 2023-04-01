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



subject = 'หลักสูตรฝึกอบรมสำหรับทำแผนอบรมประจำปี 2023'#!ต้องเปลี่ยนนะ
files =[]#!ต้องเปลี่ยนนะ
# hr_email = 'testdata.xlsx'
hr_email = 'C:/Users/USER/Downloads/qa_email_Nov_1.xlsx'#!ต้องเปลี่ยนนะ
contentori =''' เรียน  สมาชิก  ครับ<br>
ท่านที่ทำระบบ ISO9001:2015, ISO14001:2015, <br>
ISO45001:2018, IATF16949:2016, ISO50001:2018, <br>
Sedex, มรท., และ ILS<br>
===========================<br>
องค์กรใด กำลังทำแผน อบรมประจำปี 2023-2024<br>
ทั้งแบบ online/onsite<br>
วิทยากรและทีมงานQ&A "ช่วยท่านได้นะครับ"<br>
===========================<br>
อย่างน้อย  ท่าน"ต้อง" อบรมหลักสูตรต่างๆ  ดังต่อไปนี้<br>

หาก"อบรม"ไม่ครบท่านจะทำระบบไม่สมบูรณ์และจะเจอNCเรื่อยๆ<br>
ทุกๆรอบการตรวจประเมินจาก CB<br>
<br>
1.  ISO9001:2015<br>
1.1 ข้อกำหนดและการประยุกต์ใช้ISO9001:2015<br>
1.2 การปฏิบัติการเพื่อดำเนินการกับความเสี่งและโอกาส<br>
1.3 Internal Audit (ISO19011)<br>
<br>
2.  ISO14001:2015<br>
2.1 ข้อกำหนดและการประยุกต์ใช้ISO14001:2015<br>
2.2การประเมินAspectและlife cycle<br>
2.3การประเมินความเสี่ยงและโอกาส ด้านสิ่งแวดล้อม<br>
2.4 Internal Audit (ISO19011:2018)<br>
<br>
3.  ISO45001:2018<br>
3.1ข้อกำหนดและการประยุกต์ใช้ISO45001:2018<br>
3.2การประเมินความเสี่ยงและโอกาสด้านOH&S<br>
3.3 การชี้บ่งอันตรายและประเมินความเสี่ยง<br>
3.4 Internal Audit (ISO19011:2018)<br>
<br>
4.  IATF16949:2016<br>
4.1ข้อกำหดและการประยุกต์ใช้IATF16949<br>
4.2 การประมินความเสี่ยงและโอกาส<br>
4.3 Internal Audit (ISO19011:2018)<br>
4.4 Core Tools<br>
    APQP  FMEA  MSA <br>
    SPC  PPAP<br>
4.5 Special process ตามCSRsของลูกค้า<br>
    CQI8  CQI9  CQI12ฯลฯ<br>
    VDA6.3  VDA 6.5<br>
<br>
5. ISO 50001:2018<br>
5.1 ข้อกำหนดและการประยุกต์ใช้ ISO50001<br>
5.2 การประเมินความเสี่ยง ISO50001<br>
5.3  internal Audit <br>
    ISO 50001:2018 (อ้างอิง ISO19011) <br>
<br>
Note:หากท่านไม่ได้จัดอบรมให้ทีมงานครบทุกๆหน่วยงาน<br>
ระบบมันจะไม่สมบูรณ์<br>
ปัญหาจะตามมาเมื่อลูกค้า CB เข้าตรวจ<br>
===========================<br>
สนใจจัดอบรมกับQ&A<br>
Inhouse Training<br>
ติดต่อ คุณจิรภา/คุณกาญจนา<br>
027102138<br>
027102896<br>
qaquality@hotmail.com<br>
===========================<br>
องค์กรใด มีแผนจัดอบรม Inhouse Training<br>
(Online/ offline) ประจำปี 2022-2023<br>
<br>
เกี่ยวกับมาตรฐาน<br>
ISO9001:2015<br>
IATF16949:2016<br>
ISO14001:2015<br>
ISO45001:2018<br>
Core Tools For IATF<br>
APQP FMEA MSA  SPC PPAP<br>
ความเสี่ยงและโอกาส<br>
Internal Audit(ISO19011)<br>
CQI 9, CQI 8 ,CQI12<br>
8D/WHY WHY Analysis<br>
VDA  6.3   ,VDA 6.5<br>
Sedex v.6.1<br>
ISO50001:2018<br>
และอีกหลายๆหลักสูตร<br>
<br>
สนใจใช้บริการ<br>
วิทยากรจาก Q&A<br>
<br>
รบกวนจองแต่เนิ่นๆนะครับ<br>
ตอนนี้ ยังพอมีวันว่าง<br>
ในปี 2023-2024<br>
<br>
ขอบคุณทุกๆองค์กรนะครับ<br>
ที่ใช้บริการ<br>
ทั้งที่ปรึกษาและวิทยากร<br>
จากทีมงานของ Q&A<br>
===========================<br>
ทางQ&A ขอเสนอหลักสูตร<br>
ฝึกอบรมภายใน(online/offline)<br>
===========================<br>
หลักสูตร:ด้าน Safety<br>
===========================<br>
1.ความปลอดภัยในการทำงาน กับสารเคมี และการตอบโต้กรณีเกิดเหตุฉุกเฉิน<br><br>
2. การหยั่งรู้ อันตราย ป้องกันอุบัติเหตุ  KYT<br><br>
3.การสร้างพฤติกรรม ความปลอดภัยด้วยBSS(Behavior Based Safety)<br><br>
4.การวิเคราะห์งานเพื่อความปลดภัย และประเมินความเสี่ยง(JSA)<br><br>
5.การปฐมพยาบาลและการช่วยเหลื่อ ขั้นพื้นฐาน<br><br>
6.ความปลอดภัยในสถานที่ อับอากาศ 4 ผู้<br><br>
7.ความปลอดภัยในการทำงานในที่ สูง<br><br>
8. ผู้บังคับ ผูยึดเกาะและผู้ให้สัญญาณ ผู้ควบคุมปั้นจั่น  4 ผู้ (Overhead crane)<br><br>
9.เทคนิดการขับรถยก อย่างปลอดภัยและการบำรุงรักษา<br><br>
10.ความปลอดภัยในการทำงานกับลิฟต์ขนส่งสินค้า<br><br>
11. การป้องกันอันตรายจากเสียงดัง โครงการอนุรักษ์การได้ยิน<br><br>
12.Lock out & Tag out เพื่อป้องกัน อุบัติเหตุ<br><br>
13.ความปลอดภัยในการทำงานกับไฟฟ้า<br><br>
14.การยศาสตร์ เพื่อเพิ่มผลการผลิต<br><br>
15. การควบคุม ผู้รับเหมา เพื่อป้องกันอันตราย ( work permit)<br><br>
16. เจ้าหน้าที่ความปลอดภัย ระดับ<br>
    - หัวหน้างาน<br>
    - ระดับบริหาร<br>
    - คณะกรรมการความปลอดภัย(คปอ.)<br><br>
17.การอบรม ดับเพลิงขั้นต้น<br><br>
18 การดับเพลิงขั้นต้น และ การซ้อม อพยพ หนีไฟ<br><br>
19. ตรวจวัดสิ่งแวดล้อม<br>
    แสง เสียง ความร้อน ฝุ่น ปล่อง น้ำทิ้ง ฯลฯ<br><br>
===========================<br>
1. การสร้างระบบการผลิตแบบลีนในองค์กร (Lean manufacturing system) (2 วัน)<br><br>
2. การลดความสูญเปล่า 7 + 1 ประการ ในการทำงาน (7 + 1 wastes reduction)<br><br>
3. การปรับปรุงงานและเพิ่มผลผลิตด้วยเทคนิควิศวกรรมอุตสาหการ (Industrial engineering techniques)<br><br>
4. การปรับปรุงและพัฒนางานด้วยกิจกรรม 5ส (5S activity)<br><br>
5. การดำเนินกิจกรรมกลุ่มควบคุมคุณภาพในองค์กร (QCC activity)<br><br>
6. กระบวนการวิเคราะห์และแก้ไขปัญหาด้วย Root cause analysis<br><br>
7. กระบวนการวิเคราะห์และแก้ไขปัญหาด้วย Why-why analysis & 5Gen<br><br>
8. กระบวนการรายงานและแก้ไขปัญหาด้วย 8D report & Problem solving<br><br>
9. การแก้ไขปัญหาอย่างสร้างสรรค์และมีประสิทธิภาพ (Effective & Creative problem solving)<br><br>
10. การบริหารจัดการสายธารคุณค่า (Value stream management)<br><br>
11. การสร้างงานที่เป็นมาตรฐาน (Standardized work)<br><br>
12. การแก้ปัญหาและการตัดสินใจอย่างเป็นระบบ (Problem solving & Decision making as systematics)<br><br>
13. การสร้างจิตสำนึกคุณภาพภายในองค์กร (Quality awareness building)<br><br>
14. การประยุกต์ใช้เครื่องมือควบคุมคุณภาพ 7 ชนิด (QC 7 tools)<br><br>
15. การปรับปรุงพัฒนางานด้วยเทคนิคการจัดการด้วยสายตา (Visual Management)<br><br>
===========================<br>
หลักสูตร ด้านEngineering<br>
===========================<br>
1. ระบบการผลิตแบบ Lean<br><br>
2. กระบวนการแก้ไขปัญหาและรายงานปัญหาด้วย 8D<br><br>
3. กระบวนการแก้ไขปัญหาด้วย Root Cause Analysis<br><br>
4. กระบวนการวิเคราะห์ปัญหาด้วยเทคนิค Why Why Analysis<br><br>
5. กระบวนการแก้ไขปัญหาและรายงานปัญหาด้วย A3<br><br>
6. การดำเนินกิจกรรม QCC ในองค์กร<br><br>
7. การดำเนินกิจกรรม Kaizen เพื่อพัฒนาองค์กร<br><br>
8. การลดความสูญเปล่า 7+1 ประการ<br><br>
9. การดำเนินกิจกรรม 5ส ภายในองค์กร<br><br>
10. การปรับปรุงพื้นที่การทำงานด้วย Visual Management<br><br>
11. การเพิ่มผลผลิตด้วยเทคนิค Industrial Engineering<br><br>
12. การปรับปรุงพัฒนาการทำงานเพื่อเพิ่ม Productivity<br><br>
===========================<br>
หลักสูตร Human Resource<br>
===========================<br>
1. เทคนิคการสื่อสารอย่างมีประสิทธิภาพสำหรับหัวหน้างาน<br><br>
2. เทคนิคการมอบหมายงานอย่างมีประสิทธิผลสำหรับหัวหน้างาน<br><br>
3. การบริหารการทำงานอย่างมีประสิทธิภาพสำหรับหัวหน้างาน<br><br>
4. การสื่อสารอย่างมีประสิทธิภาพด้วยหลักการ Ho Ren So<br><br>
5. การพัฒนาศักยภาพการสอนงาน Train The Trainer<br><br>
6. การสอบงานให้สัมฤทธิ์ผล Effective Job Coaching<br><br>
7. การกำหนด KPI & Target และการติดตามอย่างมีประสิทธิภาพ<br><br>
8. การพัฒนาศักยภาพเพื่อเป็นหัวหน้างานที่ยอดเยี่ยม<br><br>
9. การพัฒนาศักยภาพของผู้นำที่ยอดเยี่ยมในองค์กร Excellent Leadership<br><br>
===========================<br>
1. ทักษะการสื่อสารอย่างมีประสิทธิภาพสำหรับหัวหน้างาน(Effective Communication skill for supervisors)<br><br>
2. การสื่อสารอย่างมีประสิทธิภาพด้วยหลักการ HO REN SO(Effective Communication by HO REN SO)<br><br>
3. การมอบหมายงานอย่างมีประสิทธิภาพ (Effective Job assignment)<br><br>
4. การสอนงานให้สัมฤทธิ์ผล (Effective Job coaching)<br><br>
5. การบริหารผลการปฏิบัติงานสำหรับหัวหน้างาน (Performance management system for supervisor)<br><br>
6. การเป็นหัวหน้างานที่ดี (The best supervisors)<br><br>
7. การทำงานให้สัมฤทธิ์ผลด้วยหลักการ PDCA (Effective working with PDCA concept)<br><br>
===========================<br>
1. การสร้างระบบการผลิตแบบลีนในองค์กร (Lean manufacturing system) (2 วัน)<br><br>
2. การลดความสูญเปล่า 7 + 1 ประการ ในการทำงาน (7 + 1 wastes reduction)<br><br>
3. การปรับปรุงงานและเพิ่มผลผลิตด้วยเทคนิควิศวกรรมอุตสาหการ (Industrial engineering techniques)<br><br>
4. การปรับปรุงและพัฒนางานด้วยกิจกรรม 5ส (5S activity)<br><br>
5. การดำเนินกิจกรรมกลุ่มควบคุมคุณภาพในองค์กร (QCC activity)<br><br>
6. กระบวนการวิเคราะห์และแก้ไขปัญหาด้วย Root cause analysis<br><br>
7. กระบวนการวิเคราะห์และแก้ไขปัญหาด้วย Why-why analysis & 5Gen<br><br>
8. กระบวนการรายงานและแก้ไขปัญหาด้วย 8D report & Problem solving<br><br>
9. การแก้ไขปัญหาอย่างสร้างสรรค์และมีประสิทธิภาพ (Effective & Creative problem solving)<br><br>
10. การบริหารจัดการสายธารคุณค่า (Value stream management)<br><br>
11. การสร้างงานที่เป็นมาตรฐาน (Standardized work)<br><br>
12. การแก้ปัญหาและการตัดสินใจอย่างเป็นระบบ (Problem solving & Decision making as systematics)<br><br>
13. การสร้างจิตสำนึกคุณภาพภายในองค์กร (Quality awareness building)<br><br>
14. การประยุกต์ใช้เครื่องมือควบคุมคุณภาพ 7 ชนิด (QC 7 tools)<br><br>
15. การปรับปรุงพัฒนางานด้วยเทคนิคการจัดการด้วยสายตา (Visual Management)<br><br>
===========================<br>
ติดต่อ  ฝึกอบรมได้ที่<br>
คุณ จิรภา /คุณ กาญจนา<br>
027102138  027102896<br>
0858038380<br>
qaquality@hotmail.com<br>
===========================<br>
'''
#!ต้องเปลี่ยนนะ
def sendemail(df,sender_email,password,subject,files,i,content,sheetname):
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
            print(f"sent {i+1} {sheetname}")

        except:
            pass
        server.quit()  


def sendingmarket(testdata,senderemail,passs,subject,files,sheetname,no=0):

    df = pd.read_excel(testdata,sheet_name=sheetname)
    sender_email = senderemail
    password = passs
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
        if i <100:
            continue    
        sendemail(df,sender_email,password,subject,files,i,contentori,sheetname)
#!online Zoom 45 times
#!HR 12 times
# th0 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[0],passss[0],subject,files,'Sheet1'])
# th1 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[1],passss[1],subject,files,'Sheet2'])
# th2 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[2],passss[2],subject,files,'Sheet3'])
# th3 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[3],passss[3],subject,files,'Sheet4'])
th4 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[4],passss[4],subject,files,'Sheet5'])
# th5 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[5],passss[5],subject,files,'online Zoom',0])
# th6 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[6],passss[6],subject,files,'Sheet6',0])
# th7 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[7],passss[7],subject,files,'Sheet7'])
# th8 = threading.Thread(target=sendingmarket,args=[hr_email,email_list[8],passss[8],subject,files,'Sheet9'])


# th0.start()
# th1.start()
# th2.start()
# th3.start()
th4.start()
# th5.start()
# th6.start()
# th7.start()
# th8.start()