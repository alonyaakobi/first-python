# -*- coding: utf-8 -*-
"""
Created on Fri Mar 16 13:43:04 2018

@author: ayaakobi
"""

from tkinter import *

root = Tk()

height = 5
width = 5

for i in range(height): #Rows
    for j in range(width): #Columns
        b = Entry(root, text="")
        b.grid(row=i, column=j)


for r in range(3):
   for c in range(4):
      Label(root, text='R%s/C%s'%(r,c),
         borderwidth=1 ).grid(row=r,column=c)

mainloop()




from tkinter import *

for i in range(5):
    for j in range(4):
        l = Label(text='%d.%d' % (i, j), relief=RIDGE)
        l.grid(row=i, column=j, sticky=NSEW)

mainloop()