# -*- coding: utf-8 -*-
"""
Created on Wed Mar  7 15:05:34 2018

@author: ayaakobi
"""

import threading
import time
arg1=0
arg2=1
etc=2

# your function that takes a while.
# Note: If your function returns something or if you want to pass variables in/out,
# you have to use Queues
def yourFunction(arg1,arg2,etc):
    print('111')
    time.sleep(3) #your code would replace this
    print('111')

# Setup the thread
processthread=threading.Thread(target=yourFunction,args=(arg1,arg1,etc)) #set the target function and any arguments to pass
processthread.daemon=True
processthread.start() # start the thread

#loop to check thread and display elapsed time
while processthread.isAlive():
    print (time.clock())
    time.sleep(1) # you probably want to only print every so often (i.e. every second)

print( 'Done')

