# -*- coding: utf-8 -*-
"""
Created on Sun Jun 13 15:38:40 2021

@author: yama4971
"""
import win32com.client as win32
import pandas as pd


def createReply(teacher_name, student_name, chapter, score, stu_email, email:object):
        reply = email.Reply()
        reply.to = stu_email
        if chapter == "Interlude F - MAPS":
            if score >= 70:
                newBody = f'Dear {student_name},'+ """
                <div>
                    <p>Thank you for the hand in of 
                    the first chapter assignment: “Interlude F - Maps”. I’ve now noted 
                    that you’ve started on the course. As you’ve probably seen already 
                    you managed well above the limit for pass (>70 %), well done. 
                    You’ll find the attempt, together with results and possible feedback 
                    at the Assignment’s site on the course platform (see direct link below).
                    </p>
                    <p>Please return if you have any questions!</p>
                    <p>Kind regards,</p>    
                </div>
                    """ + f'{teacher_name}' + """<p>Teacher and course admin.</p>"""
            else:
                newBody = f'Dear {student_name},'+ """
                <div>
                    <p>Thank you for the hand in of 
                    the first chapter assignment: “Interlude F”. 
                    I’ve now reviewed it and as you might have seen already 
                    it did not reach the limit for pass (≥70%). 
                    You therefore need to do a retake. This can be 
                    done one week (7 days) after your previous attempt. 
                    In the meantime you can naturally continue working on 
                    upcoming chapters!
                    </p>
                    <p>
                    You’ll find your attempt with results and possible 
                    feedback at the Assignment’s site on the course platform 
                    (see direct link below). Please make good use of this, 
                    and the course literature, when preparing for the next attempt. 
                    Return to us with any questions. Good luck!
                    </p>
                    <p>Best wishes,</p>
                </div>
                """ + f'{teacher_name}' + """<p>Teacher and course admin.</p>"""
        else:
            if score >= 70:
                newBody = f'Dear {student_name},'+ """
                <div>
                    <p>Your assignment has now been graded and you passed it (≥70%), good work!
                    </p>
                    <p>Feedback and results are available at the Assignment’s site on the course platform.
                    </p>
                    <p>Kind regards,</p>    
                </div>
                    """ + f'{teacher_name}' + """<p>Teacher and course admin.</p>"""
            else:
                newBody = f'Dear {student_name},'+ """
                <div>
                    <p>Your assignment has now been graded and unfortunately you did 
                    not reach the limit for pass (≥70%) and need to do a retake. 
                    It will be opened one week (7 days) after your last attempt. 
                    </p>
                    <p>
                    Feedback and results are available at the Assignment’s site 
                    on the course platform. Please make good use of this, 
                    and the course literature, when preparing for the next 
                    attempt and return to us with any questions. 
                    Good luck!
                    </p>
                    <p>Best wishes,</p>
                </div>
                """ + f'{teacher_name}' + """<p>Teacher and course admin.</p>"""
        reply.HTMLBody = newBody + reply.HTMLBody
        reply.Send()

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
acc = outlook.Folders.Item(2)      #Tulles folder
inbox1 = acc.folders("Inkorgen")   
inbox = inbox1.Folders.Item(4)  #assignment inbox
      

df = pd.read_excel (r'Tellus1.xlsx', sheet_name='Sheet1')
df_list1 = pd.read_excel (r'current_student.xlsx')
df_list2 = pd.read_excel (r'ST21.xlsx')
for i in range(len(df.index)):
    score = df.loc[i,'score']
    student_name = df.loc[i,'name']
    teacher_name = df.loc[i,'teacher']
    chapter = df.loc[i,'chapter']
    if df_list1[df_list1["full_name"]==student_name].empty:
        if df_list2[df_list2["full_name"]==student_name].empty:
            print ('-------')
            print ('cannot find student ' + student_name + ' in the list')
            print ('-------')
            continue
        else:
            stu_email=df_list2.loc[df_list2['full_name'] == student_name, 'email'].values[0]
    else:
        stu_email=df_list1.loc[df_list1['full_name'] == student_name, 'email'].values[0]
        
    for mailItem in inbox.Items:
            if mailItem.Subject == student_name + " has completed Assignment - " + chapter:
                print(student_name + ' sent for '+chapter )
                createReply(teacher_name, student_name, chapter, score, stu_email, mailItem)
                break
    else:
        print('--------')
        print('cannot find email of ' + student_name + ' for '+ chapter)
        print('--------')

