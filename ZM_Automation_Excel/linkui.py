from tkinter import *
from tkinter import messagebox
from linklogic import *

# initialize of tk
root = Tk()
# title
root.title("Punit's Link Automation Program")
# giving Icon
root.iconphoto(False, PhotoImage(file="Lilhakr.png"))
# background Image
bg = PhotoImage(file="Lilhakr_bg.png")
label = Label(root, image=bg)
label.place(x=0, y=0)

# Sizing of tkinter
root.geometry("400x260")

# top text - Do you want visuals?
var1 = StringVar()
label1 = Label(root, textvariable=var1, relief=FLAT)
var1.set("Do you want visuals?")
label1.pack(anchor=E, padx=30)


# visuals selection
def sel():
    selection = int(var2.get())
    if selection == 1:
        label2.config(text="You selected Yes")
    else:
        label2.config(text="You selected No")
    return selection


var2 = IntVar()

R1 = Radiobutton(root, text="Yes", variable=var2, value=1, command=sel)
R1.pack(anchor=E, padx=50)

R2 = Radiobutton(root, text="No", variable=var2, value=2, command=sel)
R2.pack(anchor=E, padx=50)

label2 = Label(root)
label2.pack(anchor=E, padx=30)

# top text - Do you want visuals?
var3 = StringVar()
label3 = Label(root, textvariable=var3, relief=FLAT)
var3.set("Input the Data here")
label3.pack(anchor=E, padx=30)


# button to enter the data
def click():
    msg = messagebox.showinfo("Opening Excel", "Please Enter the data")
    open_sheets()


B1 = Button(root, text="Input Me", command=click)
B1.place(x=290, y=120)

# top text - Do you want visuals?
var4 = StringVar()
label4 = Label(root, textvariable=var4, relief=FLAT)
var4.set("Saved the Data?\nCLick--->RUN")
label4.pack()
label4.place(x=270, y=160)


# noinspection PyArgumentList
def run():
    if int(var2.get()) == 2:
        # ChromeWebdriver.headless(login(logic()))
        # login(ChromeWebdriver.headless())
        logic(headless())

        # headless()


    else:
        # ChromeWebdriver.normal(login(logic()))
        # login(ChromeWebdriver.normal())
        logic(normal())
        # normal()
        pass


B2 = Button(root, text="Run", command=run)
B2.place(x=300, y=210)

root.mainloop()
