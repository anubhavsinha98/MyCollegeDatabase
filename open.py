#gui functionality of the project
from tkinter import *
root = Tk()
#home of the application 
var = StringVar()
#defining objects of the required tkinter classes
label = Label(root, textvariable = var, font='Helvetica 20 bold italic', fg='blue', bg='white')
var.set("Jaypee Institute Of Infromation Technology")

menu = Menu(root)
root.config(menu=menu)

subMenu = Menu(menu)
menu.add_cascade(label='Select The Hub', menu=subMenu)
subMenu.add_command(label='Robotics')
subMenu.add_command(label='Graphicas')
subMenu.add_command(label='Thespian')
#packing the different gui objects together
label.pack()
root.mainloop()
