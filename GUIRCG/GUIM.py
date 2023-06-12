# Import Tkinter 
from tkinter import *
import os
import openpyxl

# creating Main window 
r = Tk()
r.title('Report Card')
r.geometry('500x300')
r.config(bg ='#FCEEF6')


# Adding Heading in Main Window:-
Label(r,text='REPORT CARD',font=("Times New Roman", "20", "bold"), 
                    bg='purple').pack(fill=X)


# Adding Lables in Main Window
n = Label(r,text='Name:',font = ("Times New Roman", "11", "bold"), 
                bg = "#FCEEF6", fg = "black")
n.place(x = 20, y = 45,)

rn = Label(r,text='Roll Number:',font = ("Times New Roman", "11", "bold"), 
                bg = "#FCEEF6", fg = "black")
rn.place(x = 240, y = 45)

hin = Label(r,text='Hindi:',font = ("Times New Roman", "11", "bold"), 
                bg = "#FCEEF6", fg = "black")
hin.place(x = 20, y = 80)

eng = Label(r,text='English:',font = ("Times New Roman", "11", "bold"), 
                bg = "#FCEEF6", fg = "black")
eng.place(x = 20, y = 110)

sci = Label(r,text='Science:',font = ("Times New Roman", "11", "bold"), 
                bg = "#FCEEF6", fg = "black")
sci.place(x = 20, y = 140)

math = Label(r,text='Mathematics:',font = ("Times New Roman", "11", "bold"), 
                bg = "#FCEEF6", fg = "black")
math.place(x = 20, y = 170)

ss = Label(r,text='Social Science:',font = ("Times New Roman", "11", "bold"), 
                bg = "#FCEEF6", fg = "black")
ss.place(x = 20, y = 200)


#Adding Entry box into main window
e_n = Entry(font = ("Times New Roman", "11"), bg = "#FFFFFF", 
                        fg = "#5C153E", relief = GROOVE)
e_n.place(x=80, y=45)

e_rn = Entry(font = ("Times New Roman", "11"), bg = "#FFFFFF", 
                        fg = "#5C153E", relief = GROOVE)
e_rn.place(x=340, y=45)

e_hin = Entry(font = ("Times New Roman", "11"), bg = "#FFFFFF", 
                        fg = "#5C153E", relief = GROOVE)
e_hin.place(x=120, y=85)

e_eng = Entry(font = ("Times New Roman", "11"), bg = "#FFFFFF", 
                        fg = "#5C153E", relief = GROOVE)
e_eng.place(x=120, y=115)

e_sci = Entry(font = ("Times New Roman", "11"), bg = "#FFFFFF", 
                        fg = "#5C153E", relief = GROOVE)
e_sci.place(x=120, y=145)

e_math = Entry(font = ("Times New Roman", "11"), bg = "#FFFFFF", 
                        fg = "#5C153E", relief = GROOVE)
e_math.place(x=120, y=175)

e_ss = Entry(font = ("Times New Roman", "11"), bg = "#FFFFFF", 
                        fg = "#5C153E", relief = GROOVE)
e_ss.place(x=120, y=205)


# Button Function
def calculate():
    name = e_n.get()
    roll_no = int(e_rn.get())
    m_hin = int(e_hin.get())
    m_eng = int(e_eng.get())
    m_sci = int(e_sci.get())
    m_math = int(e_math.get())
    m_ss = int(e_ss.get())
    totalScore = (m_hin + m_eng + m_sci + m_math + m_ss)
    Percent = int(totalScore / 5)
    Label(r,text=totalScore,font='impack 13 bold').place(x = 390, y = 140)
    Label(r,text=Percent,font='impack 13 bold').place(x = 390, y = 200)

    Label(r,text='Total',font='impack 13 bold').place(x = 300, y = 140)
    Label(r,text='Grade',font='impack 13 bold').place(x = 300, y = 170)
    Label(r,text='Percent',font='impack 13 bold').place(x = 300, y = 200)

    if (totalScore >= 450 and totalScore <= 500):		 # Deciding grade
        grade = 'A+'
    
    elif (totalScore >= 400 and totalScore < 450):
        grade = "A"

    elif (totalScore >= 350 and totalScore < 400):
        grade = 'B'

    elif (totalScore >= 300 and totalScore < 350):
        grade = 'C'
    
    elif (totalScore >= 200 and totalScore < 300):
        grade = 'D'

    elif (totalScore >= 150 and totalScore < 200):
        grade = 'E'

    else:
        grade = 'F'	# if (totalScore >= 0 && totalScore < 200):

    Label(r,text=grade,font='impack 13 bold').place(x = 390, y = 170)

    

def enter_data():
    name = e_n.get()
    roll_no = int(e_rn.get())
    m_hin = int(e_hin.get())
    m_eng = int(e_eng.get())
    m_sci = int(e_sci.get())
    m_math = int(e_math.get())
    m_ss = int(e_ss.get())
    totalScore = (m_hin + m_eng + m_sci + m_math + m_ss)
    Percent = int(totalScore / 5)

    if (totalScore >= 450 and totalScore <= 500):		 # Deciding grade
        grade = "A+"
    
    elif (totalScore >= 400 and totalScore < 450):
        grade = "A"

    elif (totalScore >= 350 and totalScore < 400):
        grade = 'B'

    elif (totalScore >= 300 and totalScore < 350):
        grade = 'C'
    
    elif (totalScore >= 200 and totalScore < 300):
        grade = 'D'

    elif (totalScore >= 150 and totalScore < 200):
        grade = 'E'

    else:
        grade = 'F'	# if (totalScore >= 0 && totalScore < 200):

    print("==========================================")
    print("Name: ", name)
    print("Roll Number: ", roll_no)
    print("Hindi: ", m_hin)
    print("English: ", m_eng)
    print("Science: ", m_sci)
    print("Mathematics: ", m_math)
    print("Social Science: ", m_ss)
    print("................................")
    print("Total Number: ", totalScore)
    print("Percent: ", Percent)
    print("Grade: ", grade)
    print("................................")
    print("==========================================")

    filepath = "E:\Visual Studio\Project\GUIRCG\\Ayush.xlsx"

    if not os.path.exists(filepath):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        heading = ["Name", "Roll.no", "Hindi", "English","Science",
                    "Mathematics", "Social Science", "Total Number", "Percentage", "Grade"]
        sheet.append(heading)
        workbook.save(filepath)
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    sheet.append([name, roll_no, m_hin, m_eng, m_sci, m_math, m_ss, totalScore, Percent, grade])
    workbook.save(filepath)



# Adding Buttons into Main Window
button1 = Button(r,text='Calculate', bg='#6EF5FA',fg='black',
                    width=8,command=calculate)
button1.place(x=20, y =250)

button2 = Button(r,text='Exit', bg='Red',fg='black',
                    width=8,command=lambda:exit())
button2.place(x=100, y =250)

button3 = Button(r,text='Insert', bg='Green',fg='black',
                    width=8,command=enter_data)
button3.place(x=400, y =250)

mainloop()