



from tkinter import *

class Application(Frame):

    def say_hi(self):

        runfile('C:/Users/ayaakobi/.spyder-py3/limiters.py', wdir='C:/Users/ayaakobi/.spyder-py3')
        runfile('C:/Users/ayaakobi/.spyder-py3/limiters_mor_por_start_metro.py', wdir='C:/Users/ayaakobi/.spyder-py3')
        runfile('C:/Users/ayaakobi/.spyder-py3/Yview_RECL.py', wdir='C:/Users/ayaakobi/.spyder-py3')
        runfile('C:/Users/ayaakobi/.spyder-py3/areqs_factor.py', wdir='C:/Users/ayaakobi/.spyder-py3')
        runfile('C:/Users/ayaakobi/.spyder-py3/areq_WSA_table.py', wdir='C:/Users/ayaakobi/.spyder-py3')
        runfile('C:/Users/ayaakobi/.spyder-py3/araq_final.py', wdir='C:/Users/ayaakobi/.spyder-py3')
        
        print ("hi there, everyone!")

    def createWidgets(self):

        self.labal1 = Label(self)
        self.labal1["text"] = "This is Areq System"
        self.labal1.pack()

        self.QUIT = Button(self)
        self.QUIT["text"] = "QUIT"
        self.QUIT["fg"]   = "red"
        self.QUIT["command"] =  self.quit
        self.QUIT.pack({"side": "left"})

        self.hi_there = Button(self)
        self.hi_there["text"] = "Active_all_scipts",
        self.hi_there["command"] = self.say_hi
        self.hi_there.pack({"side": "left"})



    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.createWidgets()
        self.pack()

root = Tk()
app = Application(master=root)
app.mainloop()
moduleName = input('Enter module name:')
