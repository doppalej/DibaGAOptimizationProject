#!/usr/bin/env python
# coding: utf-8

# In[10]:


from DIBAUNIVERSALCALC import MAIN
import xlwings as xw
import pandas as pd
df = pd.read_excel('ORDERS.xlsm')
tk = pd.read_csv('TAKT.csv',sep = ",")

class Order:
    def __init__(self,order_id):
        self.ID = order_id
        
        for index, value in enumerate(df['Order ID']):
            if value == order_id:
                x = index     
                
        self.num_pieces = df['# Pieces'].iloc[x]
        y = df['Takt ID list'].iloc[x].split(',')
        d = []
        for i in y:
            d.append(int(i))
        self.takt_ID = d
        self.ship_date = df['Ship date'].iloc[x]
        
        z = []
        for index,value in enumerate(tk['TAKT ID']):
            if value in self.takt_ID:
                z.append(tk['TIME'].iloc[index])
        self.takt_list = z
     
    def output(self):
        return self.num_pieces, self.takt_list, self.ship_date 

def Process():
    EOD = []
    for (i,k,j) in zip(df['Order ID'],df['# Pieces'],df['Ship date']):
        order = Order(i)
        np , tl, sd = order.output()
        job_time, wrkrs = MAIN(np, tl)
        EOD.append([i,k,job_time,wrkrs,j])
        EOD.sort(key = lambda x: x[2])
    return EOD
        
        
    
x = Process()
schedule = pd.DataFrame(x,columns = ['ORDER ID','NUMBER PIECES','TOTAL HOURS','MAX STAFF','SHIP DATE'])

wb = xw.Book()
sheet = wb.sheets['Sheet1']
sheet.range('A1').value = schedule
sheet.range('A1').options(pd.DataFrame, expand='table').value        
        


# In[ ]:




