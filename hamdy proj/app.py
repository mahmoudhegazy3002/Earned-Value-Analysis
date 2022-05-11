from tkinter import *
from PIL import ImageTk,  Image
from datetime import datetime 
import json,tksheet
import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression

import matplotlib.pyplot as plt
from sklearn.preprocessing import MinMaxScaler
Window =Tk()







Window.resizable(True,True)

#my_canvas= Canvas(Window, width =1000,height =667)
#my_canvas.pack(fill='both',expand = True)


f1 = Frame(Window)
f2 = Frame(Window)
f3 = Frame(Window)
f4 = Frame(Window)
fword = Frame(Window)
fnum = Frame(Window)
frecent = Frame(Window)
fabout = Frame(Window)
fchoices = Frame(Window)
ftask1 = Frame(Window)

fwordl1=0
fnuml1=0
fwordl1= Label(fword)
fwordl1.place(x=120,y=85,)
fnuml1= Label(fnum)
fnuml1.place(x=280,y=80)

img = Image.open("images/background1.jpg")
img = img.resize((1375,1220), Image.ANTIALIAS)
bg =  ImageTk.PhotoImage(img)

i1=Label(f1,image= bg)
i1.grid(row=0,column=0,sticky="nsew")
def EarnedValueManagement(PV,AC,EV,BAC,BCWR):
    global POAC, BLB
    print(type(POAC))
    SV= EV-PV
    SPI= EV/PV
    CV= EV-AC
    CPI= EV/AC
    ETC= (BAC-EV)/CPI
    if BCWR==0:
        EAC = BAC / CPI
    else:
        EAC = AC+BCWR/(0.8*CPI+0.2*SPI)
    if CV>0 and CPI>1:
        TCPI = (BAC - EV) / (BAC - AC)
    else:
        TCPI = (BAC - EV) / (EAC - AC)
    VAC= BAC-EAC
    CI= CPI*SPI

    #numerical output 
    
    

    n1 = "Planned value= "+str(PV)
    n2 = "Actual Cost= " + str(AC)
    n3 = "Percent of actual completion= " + (str(POAC+0.000000001))[:6]
    n4 = "Earned Value= " + str(EV)
    n5 = "Schedule Variance= " + str(SV)
    n6 = "Schedule Performance Index= " + str(SPI)
    n7 = "Cost Variance= " + str(CV)
    n8 = "Cost Performance Index= " + str(CPI)
    n9 = "Budget at Completion= " + str(BAC)
    n10 = "Estimate to Completion= " + str(ETC)
    n11 = "Estimate at Completion= " + str(EAC)
    n12 = "To Complete Performance Index= " + str(TCPI)
    n13 = "Variance at Completion= " + str(VAC)
    n14 = "Critical Index= "+str(CI)
    global fnuml1

    fnuml1.config(bg='black',height=16,fg="white",width=56,text= n1 +"\n"+ n2+"\n" + n3 + "\n"+ n4 +"\n"+ n5+"\n" + n6 + "\n"+n7 +"\n"+ n8+"\n" + n9 + "\n"+n10 +"\n"+ n11+"\n" + n12 + "\n"+ n13 + "\n"+n14, font="Times 18")
      
    global fwordl1
   

    fnumtitle = Label(fnum,text="EVM Analysis",fg="Black", font="budmo 24",borderwidth=5)
    fnumtitle.place(x=520,y=15)        
    
    fnumb2=Button(fnum,text='NEXT',width=17,bg="black",borderwidth=5,fg="white", font="budmo 24",command= lambda: swap(fword))
    fnumb2.place(x=460,y=540 )
    # word output 
    if SV<0 and SPI<1:
        s1 = "We are Behind The Schedule"
    elif  SV>0 and SPI>1:
        s1 = "We are Ahead The Schedule"
    else:
        s1 ="We are on The Schedule"

    if CV<0 and CPI<1:
        s2 = "We are Over Budget"
    elif  CV>0 and CPI>1:
        s2 = "We are Under Budget"
    else:
        s2 = "We are on Budget"

    if TCPI<1:
        s3 = "We are in comfortable position"
    elif  TCPI>1:
        s3 = "We have to perform with better cost performance"
    else:
        s3 = "We can continue with same performance"

    if VAC<0:
        s4 = "It will cost " +str(int(VAC))+ " more than the planned if we continue with the same performance"
    else:
        s4 = "Contractor gain, we will end the project under budget"

    s5 = "Max performance of the project= " + str(CI)
    
    #fwordl1= Label(fword,bg='black',height=16,fg="white",width=65,text= s1 +"\n"+ s2+"\n" + s3 + "\n"+ s4 +"\n"+ s5+"\n" , font="Times 12")
    #fwordl1.place(x=70,y=65,)
    fwordl1.config(bg='black',height=12,fg="white",width=70,text= s1 +"\n"+ s2+"\n" + s3 + "\n"+ s4 +"\n"+ s5+"\n" , font="Times 22")
     
    #ftextl1 = Text(fwordl1)
    #ftextl1.insert(INSERT,'Hello....')
    #ftextl1.place(x=0,y=0)
    #ftextl1.tag_config('here',foreground="white",stat)


    fwordtitle = Label(fword,text="Project Report",fg="Black", font="budmo 24",borderwidth=5)
    fwordtitle.place(x=550,y=15)        
    
    def fsave():
        
        
        
        return f1.tkraise()

    fwordb2=Button(fword,text='SAVE',width=15,bg="black",borderwidth=5,fg="white", font="budmo 22",command= lambda: fsave())
    fwordb2.place(x=530,y=540 )
    
    minSlider = Scale(fword ,bg= "black",fg= "black",font="Times",sliderlength=170,length=296,activebackground= "black",highlightbackground='black',highlightcolor="white" ,troughcolor='black',borderwidth=5,)




def back_f2():
    global HRList, THPList ,THSList,DevList,i
    HRList= [] 
    THPList = []
    THSList = []
    DevList = []
    i=1
    taskNumber=  Label(f2,text=f'task number: {i}',fg="Black", font="budmo 20")
    taskNumber.place(x=950,y=70)
    
    return f3.tkraise()

def f3_to_f2(proname,N,BCWR):
      
    
    f2b2=Button(f2,text='NEXT',bg="black",fg="white", width=15,font="budmo 22",command= lambda: next_task_submit(Nint.get(),HRint.get(), THPint.get(),THSint.get(),deviationint.get(),proname,sponsortxt.get(), promanagertxt.get(),costumertxt.get()))
    f2b2.place(x=600,y=590)  
    

    

            
    if N == 0 or N<0:
        f3l10 = Label(f3,text='( N cannot be 0 or negative ) ',fg="red",font="Times 10 bold  italic")
        f3l10.place(x=700,y=140)
    if  BCWR<0:
        f3l10 = Label(f3,text='( BCWR cannot be negative ) ',fg="red",font="Times 10 bold  italic")
        f3l10.place(x=700,y=270)

                   
    if BCWR <0  or N <= 0 :    
        return f3.tkraise()    


    global i 
    i=0
    i= i+1
   
    return f2.tkraise()
HRList= [] 
THPList = []
THSList = []
DevList = []


POAC =0

def f4_to_f3(proname):
    f3l6 = Label(f3,text=proname,fg="Black", font="budmo 16",borderwidth=5)
    f3l6.grid(row=1,column=2,columnspan = 2,padx=0,pady=10) 
    return f3.tkraise()

def next_task_submit(N,HR, THP,THS,Dev,proname,sponsor,promanager,costumer):
    
    global i,POAC, BLB , TEV
    try :    
        HR = float(HR)
        THS = float(THS)
        THP = float(THP)
        Dev = float(Dev)        
    except ValueError:
        textLetters = '( HR,THP,THS,Deviation\ncannot have letters or signs ) '
        f2l10 = Label(f2,text=textLetters,fg="red",font="Times 10 bold  italic",width=25)
        f2l10.place(x=930,y=120)
        return f2.tkraise()
        
    if HR < 0:
        textHR = ' ( HR cannot be negative ) '
        f2l10 = Label(f2,text=textHR,fg="red",font="Times 10 bold  italic",width=25)
        f2l10.place(x=700,y=70)
    
    if THP < 0:
        textTHP =' ( THP cannot be negative ) ' 
        f2l11 = Label(f2,text=textTHP,fg="red",font="Times 10 bold  italic",width=25)
        f2l11.place(x=700,y=170)

    if THS <0:
        textTHS = "( THS cannot be negative ) "
        f2l12 = Label(f2,text= textTHS,fg="red",font="Times 10 bold  italic",width=25)
        f2l12.place(x=700,y=250)    



    if THS<0 or THP<0 or HR<0:
        return f2.tkraise()

    textTHP = ''
    textDev = ''
    textHR = ''
    textTHS = ''
    textLetters = ''
    f2l10 = Label(f2,text=textLetters,fg="red",font="Times 10 bold  italic",width=25,height = 2)
    f2l10.place(x=750,y=120)
    f2l10 = Label(f2,text=textHR,fg="red",font="Times 10 bold  italic",width=25)
    f2l10.place(x=700,y=70)

    f2l11 = Label(f2,text=textTHP,fg="red",font="Times 10 bold  italic",width=25)
    f2l11.place(x=700,y=130)
    
    f2l12 = Label(f2,text= textTHS,fg="red",font="Times 10 bold  italic",width=25)
    f2l12.place(x=700,y=190)    
    
    f2l13 = Label(f2,text = textDev,fg="red",font="Times 10 bold  italic",width=25)
    f2l13.place(x=700,y=250)    
    global HRList, THPList,THSList ,DevList
        
    if i ==1: 
        HRList = []
        THPList = []
        THSList = []
        DevList = []
    print(HRList)
    HRList.append(HR)
    THPList.append(THP)
    THSList.append(THS)
    DevList.append(Dev)
    print(HRList)
    i=i+1
    taskNumber=  Label(f2,text=f'task number: {i}',fg="Black", font="budmo 20")
    taskNumber.place(x=950,y=70)
    f2e1.delete(0,END)
    f2e1.insert(0,0)
    f2e2.delete(0,END)
    f2e2.insert(0,0)
    f2e3.delete(0,END)
    f2e3.insert(0,0)   
    f2e4.delete(0,END)
    f2e4.insert(0,0) 
    print(i)
    print(N)    
    if i ==N:
        f2b2=Button(f2,text='Submit',bg="black",fg="white", width= 15,font="budmo 22",command= lambda: next_task_submit(Nint.get(),HRint.get(), THPint.get(),THSint.get(),deviationint.get(),proname,sponsor,promanager,costumer))
        f2b2.place(x=600,y=590) 
        return f2.tkraise()

    if i == N+1:
        
        i=1
        taskNumber=  Label(f2,text=f'task number: {i}',fg="Black", font="budmo 20")
        taskNumber.place(x=950,y=70)

        N= int(Nint.get())
        BCWR = int(BCWRint.get())
        TPV=0  #Total PV
        TAC=0  #Total AC
        TEV=0  #Total EV
        BAC=0
        

        for i in range(0,N-1):
            print(i)
            HRReal  = HRList[i]
            THPReal = THPList[i]
            THSReal = THSList[i]
            DevReal = DevList[i]
            
            PV=0
            if THPReal<THSReal:
                PV = HRReal * THPReal
            else:
                PV = HRReal * THSReal
            AC= HRReal*THSReal
            
            
            POAC = THSReal/ (THPReal+DevReal)  #Percent of actual project completion
            BLB= THPReal*HRReal  #Baseline budget
            EV= BLB*POAC
            TPV= TPV+PV
            TAC= TAC+AC
            TEV= TEV+EV
            BAC= BAC+BLB
        if N == 1:
            HRReal  = HRList[0]
            THPReal = THPList[0]
            THSReal = THSList[0]
            DevReal = DevList[0]
            
            PV=0
            if THPReal<THSReal:
                PV = HRReal * THPReal
            else:
                PV = HRReal * THSReal
            AC= HRReal*THSReal
            
            
            POAC = THSReal/ (THPReal+DevReal)  #Percent of actual project completion
            BLB= THPReal*HRReal  #Baseline budget
            EV= BLB*POAC
            TPV= TPV+PV
            TAC= TAC+AC
            TEV= TEV+EV
            BAC= BAC+BLB            
            
        with open('data.json') as file:
            users=json.load(file)
            
        users[proname]={'Project Name':proname,'Sponsor':sponsor,'Project manager':promanager,"Costumer":costumer 
        ,'date created':datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"TPV":TPV,"TAC":TAC,"TEV":TEV,"BAC":BAC,"BCWR":BCWR}
        with open('data.json','w') as file:    
            json.dump(users,file)            
        
        EarnedValueManagement(TPV,TAC,TEV,BAC,BCWR) 



        return fnum.tkraise()
    



class Example(Frame):
    def __init__(self, master, *pargs):
        Frame.__init__(self, master, *pargs)



        self.image = Image.open("index.png")
        self.img_copy= self.image.copy()


        self.background_image = ImageTk.PhotoImage(self.image)

        self.background = Label(self, image=self.background_image)
        self.background.pack(fill=BOTH, expand=YES)
        self.background.bind('<Configure>', self._resize_image)

    def _resize_image(self,event):

        new_width = event.width
        new_height = event.height

        self.image = self.img_copy.resize((new_width, new_height))

        self.background_image = ImageTk.PhotoImage(self.image)
        self.background.configure(image =  self.background_image)


for frame in (f1,f2,f3,f4,fword,fnum,frecent,fabout,fchoices,ftask1):
    frame.grid(row=0,column=0,sticky="nsew")


f1.tkraise()
def swap(frame):
    
    frame.tkraise()



def CurSelect(event): 
    
    value=str(l1SM.get(l1SM.curselection()))
    with open('data.json') as file:
        users=json.load(file)
    
        for i in users :    
            if i in value :
                i= [i][0]
                
                TPV = users[i]["TPV"]

                TAC = users[i]["TAC"]
                TEV = users[i]["TEV"]
                BAC = users[i]['BAC']
                BCWR = users[i]['BCWR']
                EarnedValueManagement(TPV,TAC,TEV,BAC,BCWR)






b1=Button(f1,text='Start',bg="black",fg="white",borderwidth=5, font="budmo 20",width= 25,command=lambda: swap(f4))
b1.place(x= 480,y=250) 

b2=Button(f1,text='About',bg="black",borderwidth=3,fg="white",width = 20, font="budmo 17",command= lambda: swap(fabout))
b2.place(x=1050,y=590) 

b3=Button(f1,text='Recent',bg="black",fg="white",borderwidth=3,width=20 ,font="budmo 17",command= lambda: f1_to_frecent())
b3.place(x=30,y=590) 

ll1 = Label(f1,text='',bg="pink",fg="white", font="budmo 1")
ll1.grid(row=4,column=0,pady=  250, padx= 40) 


#frame 3


f3b1=Button(f3,text='Back',bg="black",fg="white",borderwidth=3,width=17 ,font="budmo 17",command= lambda: swap(fchoices))
f3b1.place(x=20,y=15)

f3l1 = Label(f3,text='N:',fg="Black", font="budmo 22",borderwidth=5)
f3l1.place(x=50,y=120)


Nint = DoubleVar()
f3e1 = Entry(f3,textvariable= Nint,fg="Black", font="budmo 22",borderwidth=5)
f3e1.place(x=300,y=120)
f3e1.delete(0,END)
f3e1.insert(0,0)
f3l2 = Label(f3,text='BCWR:',fg="Black", font="budmo 22",borderwidth=5)
f3l2.place(x=50,y=250)


BCWRint = DoubleVar()
f3e2 = Entry(f3,textvariable=BCWRint,bg = "white",fg="Black", font="budmo 22",borderwidth=5)
f3e2.place(x=300,y=250)
f3e2.delete(0,END)
f3e2.insert(0,0)

f3b2= Button(f3,text='NEXT',borderwidth=5,width=12,bg="black",fg="white", font="budmo 22",command=lambda: f3_to_f2(Pronametxt.get(),Nint.get(),BCWRint.get()))
f3b2.place(x=550,y=550) 


f3l4 = Label(f3,width =80,text='N: Number of Tasks \nBCWR: Budget cost of work remaining (if its given write it if not insert "0")',bg="black",fg="white", font="Times 17")
f3l4.place(x=120,y=450)  

f3l5 = Label(f3,width =80,text='',fg="white", font="Times 12")
f3l5.place(x=700,y=300)


#frame 2

f2b1=Button(f2,text='Back',bg="black",fg="white",borderwidth=3,width=17 ,font="budmo 17",command= lambda: back_f2())
f2b1.place(x=20,y=15) 


HRint = StringVar()
f2l1= Label(f2,text='HR:',fg="black", font="Times 24")
f2l1.place(x=40,y=70)  


f2e1= Entry(f2,fg="black",textvariable=HRint, font="Times 24", borderwidth=5)
f2e1.place(x=300,y=70)  
f2e1.delete(0,END)
f2e1.insert(0,"0")

f2l2 = Label(f2,text='THP:',fg="black", font="Times 24")
f2l2.place(x=40,y=150)

THPint =StringVar()
f2e2= Entry(f2,textvariable=THPint,fg="black", font="Times 24", borderwidth=5)
f2e2.place(x=300,y=150)  
f2e2.delete(0,END)
f2e2.insert(0,"0")

f2l3 = Label(f2,text='THS:',fg="black", font="Times 24")
f2l3.place(x=40, y=230)

THSint = StringVar()
f2e3= Entry(f2,fg="black",textvariable=THSint, font="Times 24", borderwidth=5)
f2e3.place(x=300,y= 230)  
f2e3.delete(0,END)
f2e3.insert(0,"0")


f2l5 = Label(f2,text='Deviation:',fg="black", font="Times 24")
f2l5.place(x=40,y=310)

deviationint = StringVar()
f2e4= Entry(f2,fg="black",textvariable=deviationint ,font="Times 24", borderwidth=5)
f2e4.place(x=300, y=310)  
f2e4.delete(0,END)
f2e4.insert(0,"0")

f2b2=Button(f2,text='NEXT',bg="black",fg="white",width=15, font="budmo 2",command= lambda: next_task_submit(Nint.get(),HRint.get(), THPint.get(),THSint.get(),deviationint.get(),Pronametxt.get(),sponsortxt.get(),promanagertxt.get(),costumertxt.get()))
f2b2.place(x=600,y=590) 


 

f2l7=Label(f2,text='HR: Hourly/Monthly/Yearly Rate\nTHP: Total Hours/Months/Years Planned\nTHS: Total Hours/Months/Years Spent\nDeviation: changes in plan, -ve for early completion, +ve for late completion and 0 \n     if there is no change in the plan',bg="#d0d3d4",width=75,fg="Black", font="budmo 20")
f2l7.place(x=50,y=400)
i=1
taskNumber = Label(f2,text=f'task number: {i}',fg="Black", font="budmo 20")
taskNumber.place(x=950,y=70)


#frecent
l1SM=Listbox(frecent,bg="#99a3a4",font="budmo 16",borderwidth=8,width=65,height=15)
l1SM.place(x=300,y=125)
def f1_to_frecent():

    with open('data.json') as file:
        users=json.load(file)
    
        for i in users :
            space =len(users[i])
            l1SM.insert(END,"     "+ users[i]["Project Name"] +"                    "+ users[i]["date created"])
    frecent.tkraise()

    l1SM.bind('<<ListboxSelect>>',CurSelect)
    


def frecent_to_f1():
    l1SM.delete(0,END)
    f1.tkraise()
def next_recent():
    l1SM.delete(-1,END)
    
    return fnum.tkraise()

frecentl1 = Label(frecent,text="Choose one of your recent projects",width=45,bg="black",fg="white", font="Times 18")
frecentl1.place(x=390,y=70)

frecentb1=Button(frecent,text='Back',bg="black",fg="white",borderwidth=3,width=17 ,font="budmo 17",command= lambda: frecent_to_f1())
frecentb1.place(x=20,y=15) 

frecentb2=Button(frecent,text='NEXT',width=25,bg="black",fg="white", borderwidth=5,font="budmo 18",command= lambda: next_recent())
frecentb2.place(x=525,y=550) 

#frame 4

f4b1=Button(f4,text='Back',bg="black",fg="white",borderwidth=3,width=17 ,font="budmo 17",command= lambda: swap(f1))
f4b1.place(x=20,y=15) 





#sponsor, Project manager, Project, Costumer , Updated on 



f4l1= Label(f4,text='Sponsor:',fg="black", font="Times 24")
f4l1.place(x=50,y=70)

sponsortxt= StringVar()
f4e1= Entry(f4,fg="black", font="Times 22", borderwidth=5,textvariable=sponsortxt)
f4e1.place(x=300,y=70)  

f4l2 = Label(f4,text='Project Manager:',fg="black", font="Times 24")
f4l2.place(x=50,y=150)

promanagertxt = StringVar()
f4e2= Entry(f4,fg="black", font="Times 22",textvariable = promanagertxt ,borderwidth=5)
f4e2.place(x=300,y=150)  

f4l3 = Label(f4,text='Costumer:',fg="black", font="Times 24")
f4l3.place(x=50, y=230)

costumertxt = StringVar()
f4e3= Entry(f4,fg="black",textvariable=costumertxt, font="Times 22", borderwidth=5)
f4e3.place(x=300,y= 230)  


f4l4 = Label(f4,text='Project Name:',fg="black", font="Times 24")
f4l4.place(x=50,y=310)

Pronametxt = StringVar()
f4e4= Entry(f4,fg="black",textvariable=Pronametxt ,font="Times 22", borderwidth=5)
f4e4.place(x=300, y=310)  

f4b2=Button(f4,text='NEXT',bg="black",fg="white", font="budmo 24",command= lambda: projectDetails(Pronametxt.get(),sponsortxt.get(),promanagertxt.get(),costumertxt.get()))
f4b2.place(x=400,y=410) 


def projectDetails(proname, sponsor, promanager, costumer ):
    if proname == '':
        f4l10 = Label(f4,text='( you cannot leave the project name empty ) ',fg="red",font="Times 12 bold  italic")
        f4l10.place(x=710,y=330)
    if sponsor == '':
        f4l10 = Label(f4,text='( you cannot leave the sponsor empty ) ',fg="red",font="Times 12 bold  italic")
        f4l10.place(x=710,y=72)
    if promanager == '':
        f4l10 = Label(f4,text='( you cannot leave the project manager empty ) ',fg="red",font="Times 12 bold  italic")
        f4l10.place(x=710,y=160) 
    if costumer == '':
        f4l10 = Label(f4,text='( you cannot leave the costumer empty ) ',fg="red",font="Times 12 bold  italic")
        f4l10.place(x=710,y=240)                    
    if sponsor == '' or proname == '' or promanager == '' or costumer == '':    
        return f4.tkraise()    

    f3l6 = Label(f3,text=proname,fg="Black", font="budmo 22",borderwidth=5)
    f3l6.place(x=1150,y=25)
    return fchoices.tkraise()    

 # About Frame 

faboutl1= Label(fabout,bg='black',height=20,fg="white",width=75,text= "Earned value management (EVM), earned value project management, or \nearned value performance management (EVPM) is a project management\ntechnique for measuring project performance and progress\n in an objective manner.Earned value management is a project management technique\n for measuring project performance and progress.\n It has the ability to combine measurements of the project management triangle:\n scope, time, and costs.\nIn a single integrated system, earned value management is able\n to provide accurate forecasts of project performance problems, which is\n an important contribution for project management.\nEarly EVM research showed that the areas of planning and control are significantly impacted\n by its use; and similarly, using the methodology improves both scope definition\n as well as the analysis of overall project performance.\n More recent research studies have shown that the principles of EVM are positive predictors\n of project success. Popularity of EVM has grown in recent years beyond \ngovernment contracting, a sector in which its importance continues to rise , in part because EVM\n can also surface in and help substantiate contract disputes" 
, font="Times 18")
faboutl1.place(x=140,y=100,)

faboutb1=Button(fabout,text='Back',bg="black",fg="white",borderwidth=3,width=17 ,font="budmo 17",command= lambda: swap(f1))
faboutb1.place(x=40,y=10) 
faboutl1= Label(fabout,bg='black',width= 20,fg="white",text="ABOUT" ,font="Times 24" )
faboutl1.place(x=450,y=20,)

#frame choices 

fchoicesb1=Button(fchoices,text='Back',bg="black",fg="white",borderwidth=3,width=17 ,font="budmo 17",command= lambda: swap(f4))
fchoicesb1.place(x=40,y=10) 

fchoicesl1= Label(fchoices,width= 20,bg= 'black',fg="white",text="EVM" ,font="Times 24" )
fchoicesl1.place(x=500,y=20,)

#fchoicesl1= Label(fchoices,fg="black",text="Do you want your tasks to be detailed?" ,font="Times 24" )
#fchoicesl1.place(x= 450,y=150,)

fchoicesb1=Button(fchoices,text='Detailed mode',width = 20,height = 5,bg="green",fg="white", font="Times 23",borderwidth=5,command= lambda: swap(ftask1))
fchoicesb1.place(x=720,y=250)

fchoicesb1=Button(fchoices,text='Quick mode',width = 20,height =5, bg="#bb0619",fg="white", font="Times 23",borderwidth=5,command= lambda: swap(f3))
fchoicesb1.place(x=330,y=250)


#-----------------------------------------------------------------------------
# def for task1 frame 
#a function that inserts a row in the spreadsheet
sheet1_data  = [["",0.0,0.0,0.0,'','',0,0,0.0,0.0]]
sheet2_data = [["",0,'','','']]
def insert_row_func(sheet1, sheet2):
    global sheet1_data, sheet2_data
    sheet1_data.append( ["",0.0,0.0,0.0,"","",0,0,0.0,0.0])
    sheet2_data.append(["",0,'','',''])

    return sheet2.set_sheet_data(data=sheet2_data, reset_col_positions=True, reset_row_positions=True, redraw=True, verify=False, reset_highlights=False),sheet1.set_sheet_data(data=sheet1_data, reset_col_positions=True, reset_row_positions=True, redraw=True, verify=False, reset_highlights=False)

def delete_row_func(sheet1, sheet2):
    global sheet1_data,sheet2_data
    sheet1_data = sheet1_data[:-1]
    sheet2_data = sheet2_data[:-1]

    return sheet2.set_sheet_data(data=sheet2_data, reset_col_positions=True, reset_row_positions=True, redraw=True, verify=False, reset_highlights=False),sheet1.set_sheet_data(data=sheet1_data, reset_col_positions=True, reset_row_positions=True, redraw=True, verify=False, reset_highlights=False)


def show_total_stats():
    text1 = commulation((np.array(sheet1_data)[:,8]).tolist())[-1]
    text2 = commulation((np.array(sheet1_data)[:,2]).tolist())[-1]
    text3 = commulation((np.array(sheet1_data)[:,3]).tolist())[-1]
    text4 = commulation((np.array(sheet1_data)[:,1]).tolist())[-1]

    ftask1l10.config(text= text1)
    ftask1l11.config(text= text2)
    ftask1l12.config(text= text3)
    ftask1l13.config(text= text4)

    #________________________________________________________________
#generating graph
def commulation(lista):
    sum1 = 0
    listb= []
    
    for i in lista:
        sum1+=float(i)
        listb.append(sum1)
    
    return listb

def multiple_points(lista):
    listb = []
    for i in range(int(min(lista)),int(max(lista))*10):
         
        listb.append(i/10)
    return listb

graph = ''
render2 = ''

def put_image():
    global graph_pic
    global graph, render2
    graph = Image.open('output.jpg')
    graph = graph.resize((550,400))
    render2=ImageTk.PhotoImage(graph)
    graph_pic.config(image=render2)
    

    graph_pic

def graph_generator(planned_duration, TPV,TEV ,TAC):
    #_______________________________________
    #comutating data
    TAC = commulation(TAC)
    TPV = commulation(TPV)
    TEV = commulation(TEV)
    planned_duration = commulation(planned_duration)
    #_______________________________________
    # creating multiple points 
    TAC_points = multiple_points(TAC)
    TPV_points = multiple_points(TPV)
    TEV_points = multiple_points(TEV)
    planned_duration_points = multiple_points(planned_duration)

    graph_points = []

    for i in planned_duration_points:
    
        if int(planned_duration[1])<= i:
                graph_points.append(i)
    #______________________________________
    # creating Linear Regression models
    a = np.array(planned_duration)
    b1 = TEV
    b2 = TPV
    b3 = TAC

    new_a = a[1:-1]
    data1 = pd.DataFrame({'x': a, 'y':b1,})
    data2 = pd.DataFrame({'x': a, 'y':b2,})
    data3 = pd.DataFrame({'x': a, 'y':b3,})

    #_______________________________________________ 
    # Scaling the data we have   
    scaler1 = MinMaxScaler()
    b1_scaled= scaler1.fit_transform(data1[["y"]])
    scaler2 = MinMaxScaler()
    b2_scaled= scaler2.fit_transform(data2[["y"]])
    scaler3 = MinMaxScaler()
    b3_scaled= scaler3.fit_transform(data3[["y"]])
    #_______________________________________________
    # creating Linear Regression Models
    LM1 = LinearRegression()
    LM1.fit(data1[["x"]],b1_scaled)
    LM2 = LinearRegression()
    LM2.fit(data2[["x"]],b2_scaled)
    LM3 = LinearRegression()
    LM3.fit(data3[["x"]],b3_scaled)
    #_______________________________________________
    # removing zeros and ones
    new_b1 = np.array(b1_scaled[1:-1]) # Getting rid of 0 and 1 values
    new_b2 = np.array(b2_scaled[1:-1]) # Getting rid of 0 and 1 values
    new_b3 = np.array(b3_scaled[1:-1]) # Getting rid of 0 and 1 values
    #_______________________________________________
    # creating log forms functions
    new_b1 = np.log((1 / new_b1) - 1)
    new_b2 = np.log((1 / new_b2) - 1)
    new_b3 = np.log((1 / new_b3) - 1)
    #_______________________________________________
    # creating models to use make sigmoid functions
    model1 = LinearRegression()
    model1.fit(new_a.reshape(-1, 1), new_b1.reshape(-1, 1))
    model2 = LinearRegression()
    model2.fit(new_a.reshape(-1, 1), new_b2.reshape(-1, 1))
    model3 = LinearRegression()
    model3.fit(new_a.reshape(-1, 1), new_b3.reshape(-1, 1))

    #taking beta and alpha from linear to plot values then 
    alpha1 = model1.coef_[0, 0]
    beta1 = model1.predict([[0]])[0, 0]
    alpha2 = model2.coef_[0, 0]
    beta2 = model2.predict([[0]])[0, 0]                              
    alpha3 = model3.coef_[0, 0]
    beta3 = model3.predict([[0]])[0, 0]  
    #_____________________________________________
    #creating new arrays
    a_points = np.array(planned_duration_points)
    graph_points = np.array(graph_points)
    #removing the first quarter of the planned duration and making it polynomial
    z_points = a_points[:int(len(a_points)/4)]
    #___________________________________________
    # creating sigmoid functions and making predictions
    # TEV
    predicted1 = 1 / (1 + np.exp(alpha1 * a_points + beta1))
    # TPV
    predicted2 = 1 / (1 + np.exp(alpha2 * a_points  + beta2))
    # TAC
    predicted3 = 1 / (1 + np.exp(alpha3 * a_points+ beta3))

    s1 = scaler1.inverse_transform(b1_scaled.reshape(-1,1))
    f1curve = scaler1.inverse_transform(predicted1.reshape(-1,1))[int(len(a_points)/4),:]
    f2curve = scaler2.inverse_transform(predicted2.reshape(-1,1))[int(len(a_points)/4),:]
    f3curve = scaler3.inverse_transform(predicted3.reshape(-1,1))[int(len(a_points)/4),:]


    model_TEV = np.poly1d(np.polyfit([0,a_points[int(len(a_points)/16)],a_points[int(len(a_points)/4)]],[0,f1curve[0]/16,f1curve[0]],5))
    model_TPV = np.poly1d(np.polyfit([0,a_points[int(len(a_points)/16)],a_points[int(len(a_points)/4)]],[0,f2curve[0]/16,f2curve[0]],5))
    model_TAC = np.poly1d(np.polyfit([0,a_points[int(len(a_points)/16)],a_points[int(len(a_points)/4)]],[0,f3curve[0]/16,f3curve[0]],5))
    # TEV
    predicted1 = 1 / (1 + np.exp(alpha1 * a_points + beta1))
    # TPV
    predicted2 = 1 / (1 + np.exp(alpha2 * a_points  + beta2))
    # TAC
    predicted3 = 1 / (1 + np.exp(alpha3 * a_points+ beta3))

    #____________________________________________
    #plotting the graph
    plt.figure(figsize = (8,5))
    plt.xlabel("Planned Duration (days)",color = "red",alpha = 0.8)

    plt.scatter(a, scaler1.inverse_transform(b1_scaled.reshape(-1,1)))
    plt.scatter(a, scaler2.inverse_transform(b2_scaled.reshape(-1,1)))
    plt.scatter(a, scaler3.inverse_transform(b3_scaled.reshape(-1,1)))

    plt.plot(a_points[int(len(a_points)/4):], scaler1.inverse_transform(predicted1[int(len(a_points)/4):].reshape(-1,1)),label= "TEV",color="blue" )
    plt.plot(a_points[int(len(a_points)/4):], scaler2.inverse_transform(predicted2[int(len(a_points)/4):].reshape(-1,1)),label = "TPV",color="orange")
    plt.plot(a_points[int(len(a_points)/4):], scaler3.inverse_transform(predicted3[int(len(a_points)/4):].reshape(-1,1)),label="TAC",color ="green")
    plt.plot([0,a_points[int(len(a_points)/34)],a_points[int(len(a_points)/4)]],model_TEV([0,a_points[int(len(a_points)/16)],a_points[int(len(a_points)/4)]]),color = 'blue',)
    plt.plot([0,a_points[int(len(a_points)/34)],a_points[int(len(a_points)/4)]],model_TPV([0,a_points[int(len(a_points)/16)],a_points[int(len(a_points)/4)]]),color = 'orange')
    plt.plot([0,a_points[int(len(a_points)/34)],a_points[int(len(a_points)/4)]],model_TAC([0,a_points[int(len(a_points)/16)],a_points[int(len(a_points)/4)]]),color = 'green')


    plt.title("Analysis")
    plt.ylabel("(in $)",color="red")
    plt.legend()
    plt.savefig("output.jpg")


def generate_btn():
    show_total_stats()
    ##--------------------------------------------------------------
    #parameters 
    global graph_pic
    global graph, render2
    TAC_new= np.array(sheet1_data)[:,2].tolist()
    TEV_new = np.array(sheet1_data)[:,3].tolist()
    TPV_new = np.array(sheet1_data)[:,1].tolist()
    planned_duration_new = np.array(sheet1_data)[:,6].tolist()
    
    graph_generator(planned_duration_new, TPV_new,TEV_new ,TAC_new)
    put_image()
    
    
#------------------------------------------------------------------------------
# task1 frame 

#________________________________________________________________
# ftask labels
ftask1l2 = Label(ftask1,text= 'Total',font = "budmo 8")
ftask1l2.place(x=2,y=485)

ftask1l4 = Label(ftask1,text= 'PV:',font = "budmo 8")
ftask1l4.place(x=30,y=500)


ftask1l5 = Label(ftask1,text= 'EV:',font = "budmo 8")
ftask1l5.place(x=170,y=500)


ftask1l6 = Label(ftask1,text= 'AC:',font = "budmo 8")
ftask1l6.place(x=280,y=500)



ftask1l7 = Label(ftask1,text= 'Cost Variance:',font = "budmo 8")
ftask1l7.place(x=410,y=500)


ftask1l10 = Label(ftask1,text= 0.0,font = "budmo 8",justify="left")
ftask1l10.place(x=485,y=500)

ftask1l11 = Label(ftask1,text= 0.0,font = "budmo 8")
ftask1l11.place(x=305,y=500)

ftask1l12 = Label(ftask1,text= 0.0,font = "budmo 8")
ftask1l12.place(x=195,y=500)


ftask1l13 = Label(ftask1,text= 0.0,font = "budmo 8")
ftask1l13.place(x=55,y=500)

#________________________________________________________________
#ftask buttons


ftask1b1=Button(ftask1,text='Back',bg="black",fg="white",borderwidth=3,width=17 ,font="budmo 17",command= lambda: swap(fchoices))
ftask1b1.place(x=40,y=10) 

ftask1b2=Button(ftask1,text='>',width=5,bg="black",fg="white", font="budmo 14",command= lambda: sheet_switches("right"))
ftask1b2.place(x=120,y=600) 
ftask1b3=Button(ftask1,text='<',width=5,bg="black",fg="white", font="budmo 14",command= lambda: sheet_switches("left"))
ftask1b3.place(x=40,y=600) 

ftask1b4=Button(ftask1,text='Generate',width=25,bg="black",fg="white", font="budmo 14",command= lambda: generate_btn())
ftask1b4.place(x=220,y=600) 

ftask1b5=Button(ftask1,text='Delete row',width=15,bg="black",fg="white", font="budmo 14",command= lambda: delete_row_func(sheet,sheet2))
ftask1b5.place(x= 620,y=540) 

ftask1b6=Button(ftask1,text='Insert row',width=15,bg="black",fg="white", font="budmo 14",command= lambda: insert_row_func(sheet,sheet2))
ftask1b6.place(x= 620,y=600) 





##sheets

#sheet Frames
sheetFrame = Frame(ftask1)
sheet2Frame = Frame(ftask1)


for sheetF in sheetFrame, sheet2Frame:
    sheetF.place(x=5,y=80,height=400,width= 795)

#sheet1

sheetFrame.tkraise()
dataHeader1 = ["Activity","Planned\nValue","Actual\nCost","Earned\nValue","Actual\nstart date",'Actual\nfinish date','Actual\nduration','Percentage\ncomplete', "Cost\nVariance",'Schedule\nVariance']
sheet = tksheet.Sheet(sheetFrame,font="Times 8",row_index_width = 25,
headers=dataHeader1,header_bg="green",header_border_fg="blue",width= 800,
height=500,header_height = "2",header_fg="black",column_width=64,header_font="Times 8")

sheet.place(x=0,y=0)



sheet.set_sheet_data(data  =sheet1_data , reset_col_positions=True, reset_row_positions=True, redraw=True, verify=False, reset_highlights=False)

#sheet.set_sheet_data([[f"{ri+cj}" for cj in range(4)] for ri in range(2)])

# table enable choices listed below:

sheet.enable_bindings(("single_select",

                       "row_select",

                       "column_width_resize",

                       "arrowkeys",

                       "right_click_popup_menu",

                       "rc_select",

                       "rc_insert_row",

                       "rc_delete_row",

                       "copy",

                       "cut",

                       "paste",

                       "delete",

                       "undo",

                       "edit_cell",
                       'redo'))
#----------------------------------------------------------------
#sheet2
dataHeader2 = ["Activity","Duration",'Start Date','Finish Date','Predeceesor']
sheet2 = tksheet.Sheet(sheet2Frame,font="Times 10",row_index_width = 120,
headers=dataHeader2,header_bg="green",header_border_fg="blue",width= 800,
height=500,header_height = "2",header_fg="black",column_width=120,header_font="Times 10")



sheet2.set_sheet_data(data=sheet2_data, reset_col_positions=True, reset_row_positions=True, redraw=True, verify=False, reset_highlights=False )

#sheet.set_sheet_data([[f"{ri+cj}" for cj in range(4)] for ri in range(2)])

# table enable choices listed below:

sheet2.enable_bindings(("single_select",

                       "row_select",

                       "column_width_resize",

                       "arrowkeys",

                       "right_click_popup_menu",

                       "rc_select",

                       "rc_insert_row",

                       "rc_delete_row",

                       "copy",

                       "cut",

                       "paste",

                       "delete",

                       "undo",

                       "edit_cell",
                       'redo'))
sheet2.place(x=0, y=0)


#----------------------------------------------------------------
#sheet 3





#define sheet funtions
sheet_number = 1
sheet_dict={1:sheetFrame,2:sheet2Frame}
def sheet_switches(direction):
    global sheet_number
    print(sheet_number)
    if direction == 'right':
        
        sheet_number += 1
        if sheet_number>2:
            print("greater than trom")
            sheet_number= 1

    elif direction == 'left':
        sheet_number -= 1
        if sheet_number == 0:
            sheet_number = 2
    ftask1l2 = Label(ftask1,text= str(sheet_number)+'/5',width = 5,font = "budmo 12")
    ftask1l2.place(x=90,y=390)

    return sheet_dict[sheet_number].tkraise()

graph_pic=Label(ftask1,text = "")
graph_pic.place(x=800,y=80)  



Window.mainloop()


       




