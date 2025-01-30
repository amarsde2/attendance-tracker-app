"""
THIS IS SIMPLE MODULE WHICH HANDLE MAILING FUNCTIONALITY FOR OUR APP

""" 

import smtplib
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
  

staff_mails = ["your staff mail address"]
from_email = "Your Admin Email"
password = "Your password"

def oneMoreLeaveMessage(student, subject):
    email_message = f"You have only one leave left in subject {subject}"
    mail = smtplib.SMTP('smtp.gmail.com', 587, timeout=120) 
    mail.starttls() 
    mail.login(from_email, password) 
    message = MIMEMultipart() 
    message['Subject'] = 'Attendance report'
    message.attach(MIMEText(email_message, 'plain')) 
    content = message.as_string() 
    mail.sendmail(from_email, student, content) 
    mail.quit() 


def leaveQuotaExceedMail(student, subject):
    email_message = f"You have used all leave available for {subject}. if you take more leave then it cause you terminate from exam. "
    mail = smtplib.SMTP('smtp.gmail.com', 587, timeout=120) 
    mail.starttls() 
    mail.login(from_email, password) 
    message = MIMEMultipart() 
    message['Subject'] = 'Attendance report'
    message.attach(MIMEText(email_message, 'plain')) 
    content = message.as_string() 
    mail.sendmail(from_email, student, content) 
    mail.quit() 


def sendNotificationToStaff(subject, students):
    email_message = f"""Please find attendance report for {subject} \n\n 
                      following students have not suffient attendance to appear in exam {",".join(students)}. 
                    """
  
    mail = smtplib.SMTP('smtp.gmail.com', 587, timeout=120) 
    mail.starttls() 
    mail.login(from_email, password) 
    message = MIMEMultipart() 
    message['Subject'] = 'Attendance report'
    message.attach(MIMEText(email_message, 'plain')) 
    content = message.as_string() 
    mail.sendmail(from_email, staff_mails, content) 
    mail.quit() 


def notifyStaffMembers(sheet, validateSubject):
    
    sub1 = []
    sub2 = []
    sub3 = []
    sub4 = []
    sub5 = []
    sub6 = []

    total_subjects = len(validateSubject)

    for i in range(2, sheet.max_row):
        
        if total_subjects >= 1 and sheet.cell(row=i, column=3).value >= 3:
           sub1.append(sheet.cell(row=i, column=2).value)
         
        if total_subjects >= 2 and sheet.cell(row=i, column=4).value >= 3:
           sub2.append(sheet.cell(row=i, column=2).value)

        if total_subjects >= 3 and sheet.cell(row=i, column=5).value >= 3:
           sub3.append(sheet.cell(row=i, column=2).value)

        if total_subjects >= 4 and sheet.cell(row=i, column=6).value >= 3:
           sub4.append(sheet.cell(row=i, column=2).value)
        
        if total_subjects >= 5 and sheet.cell(row=i, column=7).value >= 3:
           sub5.append(sheet.cell(row=i, column=2).value)

        if total_subjects >= 6 and sheet.cell(row=i, column=8).value >= 3:
           sub6.append(sheet.cell(row=i, column=2).value)


    if len(sub1) > 0:
       sendNotificationToStaff(validateSubject[0], sub1)

    if len(sub2) > 0:
       sendNotificationToStaff(validateSubject[1], sub2)

    if len(sub3) > 0:
       sendNotificationToStaff(validateSubject[2], sub3)

    if len(sub4) > 0:
       sendNotificationToStaff(validateSubject[3], sub4)

    if len(sub5) > 0:
       sendNotificationToStaff(validateSubject[4], sub5)

    if len(sub6) > 0:
       sendNotificationToStaff(validateSubject[5], sub6)


# for staff 
def mailstaff(mail_id, msg): 
    from_id = 'crazygirlaks@gmail.com'
    pwd = 'ERAkshaya485'
    to_id = mail_id 
    message = MIMEMultipart() 
    message['Subject'] = 'Lack of attendance report'
    message.attach(MIMEText(msg, 'plain')) 
    s = smtplib.SMTP('smtp.gmail.com', 587, timeout=120) 
    s.starttls() 
    s.login(from_id, pwd) 
    content = message.as_string() 
    s.sendmail(from_id, to_id, content) 
    s.quit() 
    print('Mail Sent to staff') 