from os import stat
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import webbrowser
import inspect
import gspread
from oauth2client import client, service_account
from oauth2client.service_account import ServiceAccountCredentials
from pyasn1.type.constraint import InnerTypeConstraint
import cv2
import keyboard
import numpy as np
import pyautogui
import time
face_cascade = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
status = ('non')

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json",scope)
client = gspread.authorize(creds)

root = Tk()
root.title("Google Meet Assitant")

Label(text="GOOGLE MEET ASSISTANT",font=("Arial",20),bg='#f5ccb5').place(x=130,y=20)
#ป้อนข้อมูลชั้น
Label(text="Grade",font=20,bg='#f5ccb5').place(x=90,y=70)
choice1 = StringVar(value="1")
combo=ttk.Combobox(textvariable=choice1)
combo["values"] = ("1","2","3","4","5","6")
combo.place(x=30,y=105)
#ป้อนข้อมูลห้อง
Label(text="Room",font=20,bg='#f5ccb5').place(x=310,y=70)
choice2 = StringVar(value="1")
combo=ttk.Combobox(textvariable=choice2)
combo["values"] = ("1","2","3","4","5","6")
combo.place(x=250,y=105)
#ป้อนข้อมูลเลขที่
Label(text="Number",font=20,bg='#f5ccb5').place(x=525,y=70)
choice3 = StringVar(value="1")
combo=ttk.Combobox(textvariable=choice3)
combo["values"] = (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24)
combo.place(x=470,y=105)

txt=StringVar()
mltext=Entry(root,textvariable=txt,width=20).place(x=320,y=163)

def selectnumber():
    Label(text="Your number is : "+choice3.get()).place(x=340,y=320)
    Label(text="Your room is : "+choice2.get()).place(x=180,y=320)
    Label(text="Your grade is : "+choice1.get()).place(x=60,y=320)
    webbrowser.open_new(f"https://meet.google.com/{txt.get()}")
    sheet = client.open("Google Meet Assistant").worksheet(choice1.get()+'.'+choice2.get())
    
    num1 = int(choice3.get())
    
    while True:
        now = (time.gmtime())
        
        #เก็บภาพบริเวณมุมของกล้องเรา
        g1 = pyautogui.screenshot(region=(1616,764,4,4))
        g2 = pyautogui.screenshot(region=(1811,764,4,4))
        g3 = pyautogui.screenshot(region=(1843,867,4,4))
        g4 = pyautogui.screenshot(region=(20,200,20,20))
        g5 = pyautogui.screenshot(region=(20,800,20,20))
        g6 = pyautogui.screenshot(region=(1800,250,20,20))
        #ทำให้ภาพที่ได้นำมาคำนวณได้ด้วย numpy
        gc1 = np.array(g1)
        gc2 = np.array(g2)
        gc3 = np.array(g3)
        gc4 = np.array(g4)
        gc5 = np.array(g5)
        gc6 = np.array(g6)
        #เปลี่ยนให้เป็นค่าสีแบบ HSV
        hsv1 = cv2.cvtColor(gc1, cv2.COLOR_RGB2HSV)
        hsv2 = cv2.cvtColor(gc2, cv2.COLOR_RGB2HSV)
        hsv3 = cv2.cvtColor(gc3, cv2.COLOR_RGB2HSV)
        hsv4 = cv2.cvtColor(gc4, cv2.COLOR_RGB2HSV)
        hsv5 = cv2.cvtColor(gc5, cv2.COLOR_RGB2HSV)
        hsv6 = cv2.cvtColor(gc6, cv2.COLOR_RGB2HSV)
        #กำหนดขอบเขตของสีที่ตรวจจับ
        lowergray = np.array([0, 20, 75])
        uppergray = np.array([255, 25, 85])
        lowerBGC = np.array([0, 27, 34])
        upperBGC = np.array([255, 29, 37])
        #ตรวจจับค่าที่ที่กำหนดไว้ในบริเวณที่กำหนด
        dg1 = cv2.inRange(hsv1,lowergray,uppergray)
        dg2 = cv2.inRange(hsv2,lowergray,uppergray)
        dg3 = cv2.inRange(hsv3,lowergray,uppergray)
        dg4 = cv2.inRange(hsv4,lowerBGC,upperBGC)
        dg5 = cv2.inRange(hsv5,lowerBGC,upperBGC)
        dg6 = cv2.inRange(hsv6,lowerBGC,upperBGC)
        #นับจำนวนสีที่ตรวจจับได้
        count1 = np.sum(np.nonzero(dg1))
        count2 = np.sum(np.nonzero(dg2))
        count3 = np.sum(np.nonzero(dg3))
        count4 = np.sum(np.nonzero(dg4))
        count5 = np.sum(np.nonzero(dg5))
        count6 = np.sum(np.nonzero(dg6))

        #ถ้าไม่พบค่าสีทั้ง3ที่จะแสดงว่าเปิดกล้องแล้วและเริ่มการตรวจจับใบหน้า
        if count4 == 0 and count5 == 0 and count6 == 0 :
            status = ("Other tab")

        elif count1 == 0 and count2 == 0 and count3 == 0 and count4 != 0 and count5 != 0 and count6 != 0 :
            
            #เป็บภาพในบริเวณที่เลือกไว้
            img = pyautogui.screenshot(region=(1585, 735, 295, 165))
            time.sleep(0.5)
            #เก็บอีกภาพในอีก0.5วินาทีเพื่อนำมาเปรียบเทียบกัน
            img2 = pyautogui.screenshot(region=(1585, 735, 295, 165))
            #กำหนดเฟรมก่อนหน้าและเฟรมหลังและทำเป็นค่าที่คำนวณได้
            previous_frame= np.array(img)
            current_frame= np.array(img2)
            #เปลี่ยนให้เป็นGrayเพื่อให้ง่ายต่อการตรวจจับใบหน้า
            current_frame_gray = cv2.cvtColor(current_frame, cv2.COLOR_BGR2GRAY)
            previous_frame_gray = cv2.cvtColor(previous_frame, cv2.COLOR_BGR2GRAY)
            # ให้เฟรมหลักเป็นค่าที่คำนวณได้
            frame = np.array(img)
            #ตั่งค่าสีให้ตรวจจับได้
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            #กำหนดวตัวแปรในการตรวจับใบหน้า
            faces = face_cascade.detectMultiScale(frame, 1.1, 4)
            
            #ถ้าไม่พบใบหน้าหรือตรวจจับไม่ได้จะมีสถานะเป็นไม่พบใบหน้า
            facedetectstate = np.sum(np.nonzero(faces))
            if facedetectstate == 0:
                status = ('Away')
            #ถ้าพบจะเป็นการตรวจว่าภาพที่ได้มาในแต่ละเฟรมซำ้กันหรือไม่
            else:

                    #ทำการเทียบเฟรมก่อนและหลังว่าแตกต่างกันหรือไม่
                    frame_diff = cv2.absdiff(current_frame_gray,previous_frame_gray)
                    count4 = np.sum(np.nonzero(frame_diff))
                    #ถ้าไม่แตกต่างสถานะจะเป็นภาพซ้ำ
                    if count4 == 0:
                        status = ('Image')
                        time.sleep(1)
                        
                        #ถ้าไม่ซ้ำกันจะเป็นภาพปกติ
                    else:
                        status = ('Normal')
                        time.sleep(1)
                        
                #แสดงกรอบว่าตรวจจับส่วนไหน
            
        #ถ้ายังตรวจจับสีที่กำหนดได้อยู่จะแสดงสถานะว่าปิดกล้อง                          
        else:
            status = ('Cam off')
            time.sleep(1)
        
        # show the frame
        la = ("C" + str(num1+1))
        sheet.update(la, (status))
        time.sleep(1)

        if now[4] % 15 == 0:
            sheet.update(la, ("Away"))
            check1 = Tk()
            check1.withdraw()
            messagebox.showerror("Google Meet Assistant","Are you learning?")
            check1.destroy()
            sheet.update(la, ("Normal"))
            time.sleep(60)

        # if the user clicks q, it exits
        if keyboard.is_pressed("h"):
            sheet.update(la, ("Offline"))
            break
        
btn=Button(text="Start google meet",command=selectnumber,width="20",font="20",bg="#f8893c", activebackground='#944105', activeforeground='white').place(x=230,y=200)
Label(text="Meet Code : ",font=15,bg='#f5ccb5').place(x=180,y=160)

def exitprogram():
    root.destroy()
btn=Button(text="Close Program",command=exitprogram,width="20",font="20",bg="#f8893c", activebackground='#944105', activeforeground='white').place(x=230,y=250)
root.geometry("700x300+600+250")
root.configure(bg='#f5ccb5')
root.mainloop()


