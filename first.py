#!/usr/binp/python
import os
import time
from tkinter import *
import tkinter.messagebox
import xlrd
from xlutils.copy import copy
import xlwt
from tkinter import messagebox
from tkinter.ttk import Progressbar
from tkinter import ttk

#.......MOVING BAR........#

def baar(a,bar):
	for i in range(0,101,10):
		bar['value']=i
		time.sleep(0.05)
		a.update_idletasks()


#........FOR SUBMITTING ENTRIES........#

def enter(nameHub,root_1,top_frame,bottom_frame,bar,entry_name,entry_enroll,entry_mob,entry_email,entry_branch):
	#a=str(entry_name.get())
	#b=str(entry_enroll.get())
	#c=str(entry_mob.get())
	#d=str(entry_email.get())
	#e=str(entry_branch.get())
   # allocating name of the selected hub to loc 
	if nameHub=='IEEE':
		loc = ('IEEE.xls')
	elif nameHub=='CICE':
		loc = ('CICE.xls')
	elif nameHub=='UCR':
		loc = ('UCR.xls')
	elif nameHub=='Thespian Cirlce':
		loc = ('Thespian Circle.xls')
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
	w_sheet.write(count,0,entry_name.get())
	w_sheet.write(count,1,entry_enroll.get())
	w_sheet.write(count,2,entry_mob.get())
	w_sheet.write(count,3,entry_email.get())
	w_sheet.write(count,4,entry_branch.get())
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
		hub = 'Thespian Cirlce.xls'
	elif nameHub=='Graphicas':
		loc = ('Graphicas.xls')
		hub = 'Graphicas.xls'
	else:
		loc = ('Knuth.xls')
		hub = 'Knuth.xls'
	wb.save(hub)
	baar(bottom_frame,bar)
	recur(nameHub,root_1,top_frame,bottom_frame)

#..........FOR RESET BUTTON.........#
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

	#name_1 = entry_name.get()
	#enroll_1 = entry_enroll.get()
	#mob_1 = entry_mob.get()
	#email_1 = entry_email.get()
	#branch_1 = entry_branch.get()

	style = ttk.Style()
	style.theme_use('default')
	style.configure("blue.Horizontal.TProgressbar", background='blue')
	bar = Progressbar(bottom_frame, length=100, mode = 'determinate')

	submit_button = Button(bottom_frame,text="SUBMIT", fg="blue",command=lambda: enter(a,root_1,top_frame,bottom_frame,bar,entry_name,entry_enroll,entry_mob,entry_email,entry_branch))
	reset_button = Button(bottom_frame,text="RESET", fg="blue",command=lambda:recur(a,root_1,top_frame,bottom_frame))

	submit_button.grid(row=10,column=0,sticky="E")
	reset_button.grid(row=10,column=1)

	bar.grid(row=11,columnspan=2)

#...............MAIN GUI WINDOW................#

def fun(a,root_first):
	root_first.destroy()
	#root.destroy()
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

	#name_1 = entry_name.get()
	#enroll_1 = entry_enroll.get()
	#mob_1 = entry_mob.get()
	#email_1 = entry_email.get()
	#branch_1 = entry_branch.get()

	####>>>>>>BAR CODE<<<<<<####
	style = ttk.Style()
	style.theme_use('default')
	style.configure("blue.Horizontal.TProgressbar", background='blue')
	bar = Progressbar(bottom_frame, length=100, mode = 'determinate')


	submit_button = Button(bottom_frame,text="SUBMIT",command=lambda: enter(a,root_1,top_frame,bottom_frame,bar,entry_name,entry_enroll,entry_mob,entry_email,entry_branch),fg="blue")
	reset_button = Button(bottom_frame,text="RESET",command=lambda: recur(a,root_1,top_frame,bottom_frame), fg="blue")

	submit_button.grid(row=10,column=0,sticky="E")
	reset_button.grid(row=10,column=1)


	bar.grid(row=11,columnspan=2)

#.......FOR SEARCHING......#

def searchAnything(a,root_first,entry):
	ent = entry.get()
	if a =='IEEE':
		loc = ('IEEE.xls')
	elif a =='CICE':
		loc = ('CICE.xls')
	elif a =='UCR':
		loc = ('UCR.xls')
	elif a =='Thespian Cirlce':
		loc = ('Thespian Circle')
	elif a =='Graphicas':
		loc = ('Graphicas.xls')
	else:
		loc = ('Knuth.xls')
	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(0)
	rows = sheet.nrows
	flag=0
	for i in range(rows):
		if sheet.cell_value(i,0)==ent:
			print(sheet.cell_value(i,0))
			print(sheet.cell_value(i,1))
			print(sheet.cell_value(i,2))
			print(sheet.cell_value(i,3))
			print(sheet.cell_value(i,4))
			messagebox.showinfo(a,sheet.cell_value(i,0)+"\n"+sheet.cell_value(i,1)+"\n"+sheet.cell_value(i,2)+"\n"+sheet.cell_value(i,3)+"\n"+sheet.cell_value(i,4))
			flag=1
		elif sheet.cell_value(i,1)==ent:
			print(sheet.cell_value(i,0))
			print(sheet.cell_value(i,1))
			print(sheet.cell_value(i,2))
			print(sheet.cell_value(i,3))
			print(sheet.cell_value(i,4))
			flag=1
			messagebox.showinfo(a,sheet.cell_value(i,0)+"\n"+sheet.cell_value(i,1)+"\n"+sheet.cell_value(i,2)+"\n"+sheet.cell_value(i,3)+"\n"+sheet.cell_value(i,4))
			break
		elif sheet.cell_value(i,2)==ent:
			print(sheet.cell_value(i,0))
			print(sheet.cell_value(i,1))
			print(sheet.cell_value(i,2))
			print(sheet.cell_value(i,3))
			print(sheet.cell_value(i,4))
			flag=1
			messagebox.showinfo(a,sheet.cell_value(i,0)+"\n"+sheet.cell_value(i,1)+"\n"+sheet.cell_value(i,2)+"\n"+sheet.cell_value(i,3)+"\n"+sheet.cell_value(i,4))
			break
		elif sheet.cell_value(i,3)==ent:
			print(sheet.cell_value(i,0))
			print(sheet.cell_value(i,1))
			print(sheet.cell_value(i,2))
			print(sheet.cell_value(i,3))
			print(sheet.cell_value(i,4))
			flag=1
			messagebox.showinfo(a,sheet.cell_value(i,0)+"\n"+sheet.cell_value(i,1)+"\n"+sheet.cell_value(i,2)+"\n"+sheet.cell_value(i,3)+"\n"+sheet.cell_value(i,4))
			break
		elif sheet.cell_value(i,4)==ent:
			print(sheet.cell_value(i,0))
			print(sheet.cell_value(i,1))
			print(sheet.cell_value(i,2))
			print(sheet.cell_value(i,3))
			print(sheet.cell_value(i,4))
			flag=1
			messagebox.showinfo(a,sheet.cell_value(i,0)+"\n"+sheet.cell_value(i,1)+"\n"+sheet.cell_value(i,2)+"\n"+sheet.cell_value(i,3)+"\n"+sheet.cell_value(i,4))
			break
	if flag==0:
		print("Not Found")
		messagebox.showwarning(a,"NOT FOUND!\n {:: _ ::}")

#.........SEARCH FUNCTION..........#
def search(a,root_first):
	root_first.destroy()
	root_search = Tk()
	root_search.title(a)
	label = Label(root_search,text="Enter Anything")
	enter = Entry(root_search)
	buttonSearch =  Button(root_search,text="Search",command=lambda: searchAnything(a,root_first,enter))
	label.grid(row=0,column=0,sticky="E")
	enter.grid(row=0,column=1)
	buttonSearch.grid(row=1,column=1,sticky="W")

#.........SECOND GUI WINDOW.........#
def first(a):
	root.destroy()
	root_first = Tk()
	root_first.title(a)
	topFrame = Frame(root_first)
	topFrame.pack()
	bottomFrame = Frame(root_first)
	bottomFrame.pack()
	label = Label(topFrame,text="Welcome To "+a,font=("Algerian", 36),fg="Green",bg="White")
	label.pack()
	but1 = Button(bottomFrame,text="Submit Entries",command = lambda: fun(a,root_first))
	but2 = Button(bottomFrame,text="Search Entry",command = lambda: search(a,root_first))
	but1.grid(row=0,column=0,sticky=W+E+N+S)
	but2.grid(row=0,column=1,sticky=W+E+N+S)


########     |\/\      \    |       #######
########    |    \ AIN  \/\/ INDOW  #######
########                            #######
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

#Picture
photo = PhotoImage(file="download.png")
label1 = Label(topFrame, image=photo,justify=LEFT)
label1.pack()

#label
label2 = Label(midFrame,text="\nClick On The Hub\n",font=("Courier", 20),fg="green")
label2.pack()

#Buttons
b1=Button(bottomFrame,text="Robotics",fg="red",command=lambda: first("UCR"),height=3,width=10)
b2=Button(bottomFrame,text="CICE",fg="red",command=lambda: first("CICE"),height=3,width=10)
b3=Button(bottomFrame,text="IEEE",fg="red",command=lambda: first("IEEE"),height=3,width=10)
b4=Button(bottomFrame,text="Graphicas",fg="red",command=lambda: first("Graphicas"),height=3,width=10)
b5=Button(bottomFrame,text="Thespian Cirle",fg="red",command=lambda: first("Thespian Circle"),height=3,width=10)
b6=Button(bottomFrame,text="Knuth",fg="red",command=lambda: first("Knuth"),height=3,width=10)

b1.grid(row=0,column=0,sticky=W+E+N+S)
b2.grid(row=0,column=2,sticky=W+E+N+S)
b3.grid(row=0,column=4,sticky=W+E+N+S)
b4.grid(row=2,column=0,sticky=W+E+N+S)
b5.grid(row=2,column=2,sticky=W+E+N+S)
b6.grid(row=2,column=4,sticky=W+E+N+S)

root.mainloop()
