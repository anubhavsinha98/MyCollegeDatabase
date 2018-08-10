from tkinter import *
root = Tk()

var = StringVar()

label = Label(root, textvariable = var, font='Helvetica 20 bold italic', fg='blue', bg='white')
var.set("Jaypee Institute Of Infromation Technology")

menu = Menu(root)
root.config(menu=menu)

subMenu = Menu(menu)
menu.add_cascade(label='Select The Hub', menu=subMenu)
subMenu.add_command(label='Robotics')
subMenu.add_command(label='Graphicas')
subMenu.add_command(label='Thespian')

label.pack()
root.mainloop()
