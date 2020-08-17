from tkinter import *
mg= Tk()

def enter():
    b=a.get()
    Lab2= Label(mg,text=b,fg='red',bg='green',font=10).pack()
def delete():
    Lab3= Label(mg,text='deleting...',fg='black',bg='red',font=10).pack()
def new():
    Lab4= Label(mg,text='New file Loading',fg='black',bg='red',font=10).pack()

a=StringVar()

mg.title("First Window")
mg.geometry("500x500+200+200")

Lab1= Label(mg,text='what you wanna do',fg='red',bg='green',font=10).pack()
b1=Button(mg,text='enter',fg='blue',bg='yellow',command= enter,font=10).pack()
b2=Button(mg,text='delete',fg='blue',bg='red',command= delete,font=10).pack()


text = Entry(mg,textvariable=a).pack()



mm1=Menu()

list1=Menu()
list1.add_command(label='New',command=new)
list1.add_command(label='Save')
list1.add_command(label='Open')


list2=Menu()
list2.add_command(label='Undo')
list2.add_command(label='Redo')
list2.add_command(label='Something')

mm1.add_cascade(label="File",menu=list1)
mm1.add_cascade(label="View",menu=list2)
mm1.add_cascade(label="Edit")
mm1.add_cascade(label="Tools")
mg.config(menu=mm1)



mg.mainloop()
