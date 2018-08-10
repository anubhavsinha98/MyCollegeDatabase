#!/usr/binp/python
import os
from tkinter import *
import tkinter.messagebox
import xlrd
from xlutils.copy import copy
import xlwt


def enter(nameHub,root_1,top_frame,bottom_frame,entry_name,entry_enroll,entry_mob,entry_email,entry_branch):
	a=entry_name.get()
	b=entry_enroll.get()
	c=entry_enroll.get()
	d=entry_enroll.get()
	e=entry_branch.get()
	if nameHub=='IEEE':
		loc = ('IEEE.xls')
	elif nameHub=='CICE':
		loc = ('CICE.xls')
	elif nameHub=='UCR':
		loc = ('UCR.xls')
	elif nameHub=='Thespian Cirlce':
		loc = ('Thespian Circle')
	elif nameHub=='Graphicas':
		loc = ('Graphicas.xls')
	else:
		loc = ('Knuth.xls')
	rb = xlrd.open_workbook(loc)
	sheet = rb.sheet_by_index(0)
	sheet.cell_value(0, 0)
	count = sheet.nrows
	wb = copy(rb)
	w_sheet = wb.get_sheet(0)
	w_sheet.write(count,0,a)
	w_sheet.write(count,1,b)
	w_sheet.write(count,2,c)
	w_sheet.write(count,3,d)
	w_sheet.write(count,4,e)
	if nameHub=='IEEE':
		loc = ('IEEE.xls')
		hub = 'IEEE.xls'
	elif nameHub=='CICE':
		loc = ('CICE.xls')
		hub = 'CICE.xls'
	elif nameHub=='UCR':
		loc = ('UCR.xls')
		hub = 'UCR.xls'
	elif nameHub=='Thespian Cirlce':
		loc = ('Thespian Circle')
		hub = 'Thespian.xls'
	elif nameHub=='Graphicas':
		loc = ('Graphicas.xls')
		hub = 'Graphicas.xls'
	else:
		loc = ('Knuth.xls')
		hub = 'Knuth.xls'
	wb.save(hub)
	recur(a,root_1,top_frame,bottom_frame)
	

def recur(a,root_1,top_frame,bottom_frame):
	label_name=Label(top_frame,text="Name")
	label_enroll=Label(top_frame,text="Enroll. No.")
	label_mob=Label(top_frame,text="Contact")
	label_email=Label(top_frame,text="Email")
	label_branch=Label(top_frame,text="Branch")

	entry_name = Entry(top_frame)
	entry_enroll = Entry(top_frame)
	entry_mob = Entry(top_frame)
	entry_email = Entry(top_frame)
	entry_branch = Entry(top_frame)
	
	label_name.grid(row=0,column=0,sticky="E")
	label_enroll.grid(row=2,column=0,sticky="E")	
	label_mob.grid(row=4,column=0,sticky="E")
	label_email.grid(row=6,column=0,sticky="E")
	label_branch.grid(row=8,column=0,sticky="E")
	
	entry_name.grid(row=0,column=1)
	entry_enroll.grid(row=2,column=1)
	entry_mob.grid(row=4,column=1)
	entry_email.grid(row=6,column=1)
	entry_branch.grid(row=8,column=1)
	
	submit_button = Button(bottom_frame,text="SUBMIT", fg="blue",command=lambda: enter(a,root_1,top_frame,bottom_frame,entry_name,entry_enroll,entry_mob,entry_email,entry_branch))
	reset_button = Button(bottom_frame,text="RESET", fg="blue",command=lambda:recur(a,root_1,top_frame,bottom_frame))
	
	submit_button.grid(row=10,column=0,sticky="E")
	reset_button.grid(row=10,column=1)



def submit(a,root_1,top_frame,bottom_frame):
	print(a)
	recur(a,root_1,top_frame,bottom_frame)


def fun(a):
	root.destroy()
	root_1 = Tk()
	root_1.title(a)

	top_frame = Frame(root_1)
	top_frame.pack()
	bottom_frame = Frame(root_1)
	bottom_frame.pack()

	label_name=Label(top_frame,text="Name")
	label_enroll=Label(top_frame,text="Enroll. No.")
	label_mob=Label(top_frame,text="Contact")
	label_email=Label(top_frame,text="Email")
	label_branch=Label(top_frame,text="Branch")


	entry_name = Entry(top_frame)
	entry_enroll = Entry(top_frame)
	entry_mob = Entry(top_frame)
	entry_email = Entry(top_frame)
	entry_branch = Entry(top_frame)

	label_name.grid(row=0,column=0,sticky="E")
	label_enroll.grid(row=2,column=0,sticky="E")	
	label_mob.grid(row=4,column=0,sticky="E")
	label_email.grid(row=6,column=0,sticky="E")
	label_branch.grid(row=8,column=0,sticky="E")
	
	entry_name.grid(row=0,column=1)
	entry_enroll.grid(row=2,column=1)
	entry_mob.grid(row=4,column=1)
	entry_email.grid(row=6,column=1)
	entry_branch.grid(row=8,column=1)
	
	name_1 = entry_name.get()
	enroll_1 = entry_enroll.get()
	mob_1 = entry_mob.get()
	email_1 = entry_email.get()
	branch_1 = entry_branch.get()

	
	#print(name)
	#print(enroll_no) 
	
	submit_button = Button(bottom_frame,text="SUBMIT",command=lambda: enter(a,root_1,top_frame,bottom_frame,entry_name,entry_enroll,entry_mob,entry_email,entry_branch),fg="blue")
	reset_button = Button(bottom_frame,text="RESET",command=lambda: recur(a,root_1,top_frame,bottom_frame), fg="blue")
	
	submit_button.grid(row=10,column=0,sticky="E")
	reset_button.grid(row=10,column=1)
	


root = Tk()
root.title("JIIT")

#FRAMES SECTION
topFrame = Frame(root)
topFrame.pack()

midFrame = Frame(root)
midFrame.pack()

bottomFrame = Frame(root)
bottomFrame.pack()

#label
label = Label(topFrame,text="JAYPEE INSTITUTE OF INFORMATION TECHNOLOGY",font=("Courier", 36),fg="Blue",bg="White")
label.pack()

#Pic
photo = PhotoImage(file="download.png")
label1 = Label(topFrame, image=photo,justify=LEFT)
label1.pack()

#label
label2 = Label(midFrame,text="\nClick On The Hub\n",font=("Courier", 20),fg="green")
label2.pack()

#Buttons
b1=Button(bottomFrame,text="Robotics",fg="red",command=lambda: fun("UCR"),height=3,width=10)
b2=Button(bottomFrame,text="CICE",fg="red",command=lambda: fun("CICE"),height=3,width=10)
b3=Button(bottomFrame,text="IEEE",fg="red",command=lambda: fun("IEEE"),height=3,width=10)
b4=Button(bottomFrame,text="Graphicas",fg="red",command=lambda: fun("Graphicas"),height=3,width=10)
b5=Button(bottomFrame,text="Thespian Cirle",fg="red",command=lambda: fun("Thespian Circle"),height=3,width=10)
b6=Button(bottomFrame,text="Knuth",fg="red",command=lambda: fun("Knuth"),height=3,width=10)

b1.grid(row=0,column=0,sticky=W+E+N+S)
b2.grid(row=0,column=2,sticky=W+E+N+S)
b3.grid(row=0,column=4,sticky=W+E+N+S)
b4.grid(row=2,column=0,sticky=W+E+N+S)
b5.grid(row=2,column=2,sticky=W+E+N+S)
b6.grid(row=2,column=4,sticky=W+E+N+S)

root.mainloop()
