#!/usr/bin/env python
# coding: utf-8

# In[4]:


import xlwings as xw
import pandas as pd
import numpy as np


class Step:
    def __init__(self, takt, num_worker, inv):
        self.takt = takt
        self.counter = 0
        self.num_worker = num_worker
        self.inv = inv
        self.working = False
        self.passed = 1*self.num_worker
        
    def grab_piece(self):
        self.working = True
        if (self.inv - 1*self.num_worker) < 0:
            self.inv = 0
            self.passed = -1 * (0-1*self.num_worker)
        else:
            self.inv -= 1*self.num_worker
        
    def pass_piece(self):
        self.working = False
        self.counter = 0  
        
    def piece_needed(self):
        if self.inv <= 0:
            return False
        else:
            return True
        
    def move_workers(self):
        self.num_worker = 0
        
def MAIN(piece_num, takt_list):        
    start = piece_num
    
    steps = len(takt_list)
    staff = steps + 2
    
    takts = takt_list

    worker_line = steps * [0]
    unassigned = staff
    while unassigned != 0:
        for index,value in enumerate(worker_line):
            if unassigned > 0:
                worker_line[index] += 1
                unassigned -= 1

    line = [Step(0,0,0) for i in range(steps)]
    line[0].inv = int(start)


    for obj,wrkr in zip(line,worker_line):
        obj.num_worker = int(wrkr)


    for obj,t in zip(line,takts):
        obj.takt = int(t)


    needed = line[0].inv
    complete = 0
    clock = 0

    while complete < needed:
        for index,step in enumerate(line):
            if step.working == False:
                if step.piece_needed() == True:
                    step.grab_piece()
                    continue
            if step.working == True:
                if step.counter == step.takt:
                    step.pass_piece()
                    if index != (len(line) - 1):
                        line[index+1].inv += 1*step.num_worker
                    else:
                        complete += 1*step.num_worker
                        if complete > needed:
                            complete = needed
                else:
                    step.counter += 1

            if step.working == 'Done':
                if index != (len(line) - 1):
                        line[index+1].num_worker += step.num_worker
                        step.move_workers()

            if step.inv <= 0 and step.working == False:

                if index == 0:
                    step.working = 'Done'
                if index != 0 and line[index-1].working == 'Done':
                    step.working = 'Done'

        clock += 1
        #print(clock,line[0].inv,line[0].working,line[1].inv,line[1].working,line[2].inv,line[2].working,line[3].inv,line[3].working,complete)
        #print(clock,line[0].inv,line[0].num_worker,line[1].inv,line[1].num_worker,line[2].inv,line[2].num_worker,line[3].inv,line[3].num_worker,complete)
    
    
    perfect_time = round(clock/60/60,2)
    
    #mins
    mini_breaks = perfect_time*5
    inefficiency_factor = 1.2
    prep_time = start/10
    
   
    
    total_time_mins = perfect_time*60*inefficiency_factor + mini_breaks + prep_time
    total_time = round(total_time_mins/60,2)
    
    return total_time, staff


# In[2]:





# In[ ]:





# In[ ]:




