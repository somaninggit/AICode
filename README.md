# AICode
import csv
import datetime
import xlsxwriter
import openpyxl
import speech_recognition as sr
import pyttsx3
from datetime import date

today=date.today()
print("today's  date is : ",today)

listener=sr.Recognizer()
engine=pyttsx3.init()
voices=engine.getProperty("voices")
engine.setProperty("voice",voices[1].id)
workbook = xlsxwriter.Workbook('ATTENDANCE1.xlsx')
worksheet = workbook.add_worksheet()

l=[]


 
usn = (
    ['Skanda', '117'],
    ['Shridhar',   '114'],
    ['Tanveer',  '126'],
    ['Vinayak',    '140'],
    ['SHIVLINGESH',    '110'],
    ['SHIVRAJ KOLI',    '112'],
    ['VIKAS',    '139'],
    ['SAFFAN',    '101'],
    ['VIVEK',    '144'],
    ['VISHNU',    '141'],
    ['somu','119'],
   

) 
'''
name=["skanda","somu","Tan","shridhar"]
usn=["117","119","126","114"]
''' 
def take_com():
    try:
        with sr.Microphone() as source:
            print("listening")
            voice=listener.listen(source)
            command=listener.recognize_google(voice)
            
    except:
        pass
    
    return command

def alexarun():
    com=take_com()
    #print(com)
    if com.isnumeric():
        count = 0
        for i,j in usn:
            count = count + 1
            if(j==com):
                print(com+" is present")
                l.append(com)
                print(l)
                
                break
            if(count==len(usn)-1):
                print("enter the valid number")        


       
        
        
    elif (com=="exit"):
        print("exited")
        row = 0
        col = 0
        for k in l:
            for i,j in usn:
                if j==k:
                    worksheet.write(row, col,i)
                    worksheet.write(row, col + 1, j)
                    row=row+1
                    break
        workbook.close()
        exit() 
 
    else:
        print("please say the command again")

while(True):
    alexarun()
