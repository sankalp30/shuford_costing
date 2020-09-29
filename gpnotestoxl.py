# -*- coding: utf-8 -*-
"""
Created on Tue Jul 28 14:30:48 2020

@author: SankalpMishra
"""

import pandas as pd
import numpy as np
import pyautogui as pag
import time
#%%
"""
Left monitor is landscape, and other two screens are portrait.
Check coordinates value before running.
Check vertical adjustment of monitors relative to each other
Warning: If run with incorrect coordinate values, thisscript can damage data on great plains.
"""

path = r'\\shufordyarnsllc.local\SYDFS\UserFiles$\SankalpMishra\Desktop\Sankalp_all\cost_assistanceproject_8_19_2019\Hickory Spinners'
path = path.replace('\\', '/')
col_cno = ['cno', 'yno', 'ply', 'blend', 'tr', 'inv_description', 'case', 'units', 'tare', 'customer']

#df_cno = pd.read_excel(path + 'Cheat Sheet 7-19-20' + '.xlsx', sheet_name = 'ALL YARNS' , usecols = [0,1,2,3,4,5,6,7,8,9])
#df_cno.columns = col_cno
#df_cno['cno'] = df_cno['cno'].astype(str)

#%%

df_input = pd.read_excel(path + '/mar_2020/mar20_hs_pp' + '.xlsx', sheet_name= 'hs_pp')
df_input['cno'] = df_input['Item Number'].map(lambda x: str(x)[:-6]) #searching for full name helps with duplicate numbers in the database
#df_input['cno'] = df_input['cno'].map(lambda x: str(x)[-6:])
#%%

list_try = ['380604', '122404', '050402', '372002', '037201']

for cno in list(df_input['cno']):
#for cno in list_try:
    #item name enter.
#    c = cno[-6:]
#    pag.click(x=2596, y=256)
#    pag.typewrite(cno)
#    time.sleep(0.5)
#    #item name search using notes button
#    pag.click(x=2831, y=257)
#    time.sleep(0.5)
#    #open notes
#    pag.click(x=2831, y=257)
#    time.sleep(0.5)
#    ##notes select all, notes copy
#    pag.click(x=3742, y=531)
#    pag.keyDown('ctrl')
#    pag.keyDown('a')
#    pag.keyUp('a')
#    pag.keyUp('ctrl')
#    time.sleep(0.5)
#    
#    pag.keyDown('ctrl')
#    pag.keyDown('c')
#    pag.keyUp('c')
#    pag.keyUp('ctrl')
#    time.sleep(0.5)
#    
#    #move to excel workbook || add new sheet
#    pag.click(x=-215, y=732)
#    pag.keyDown('shift')
#    pag.keyDown('f11')
#    pag.keyUp('f11')
#    pag.keyUp('shift')
#    
#    ##move to A1 cell
#    pag.click(x=-1027, y=216)
#    time.sleep(0.5)
#    #notes paste
#    pag.keyDown('ctrl')
#    pag.keyDown('v')
#    pag.keyUp('v')
#    pag.keyUp('ctrl')
#    time.sleep(0.5)
#    
#    
#    #rename sheet
#    pag.click(x=-215, y=732)
#    pag.keyDown('alt')
#    pag.keyDown('h')
#    pag.keyUp('h')
#    pag.keyDown('o')
#    pag.keyUp('o')
#    pag.keyDown('r')
#    pag.keyUp('r')
#    pag.keyUp('alt')
#    time.sleep(0.5)
#    pag.typewrite(c)
#    time.sleep(0.5)
#    pag.click(x=-1009, y=219)
#    
#    #back to notes, close notes
#    pag.click(x=3742, y=531)
#    time.sleep(0.5)
#    pag.click(x=3798, y=149,)
#    time.sleep(0.5)
#    #item nameclear
#    pag.click(x=2530, y=226)
#    time.sleep(0.5)
    
    #####
    ############
    
    c = cno[-6:]
    pag.click(x=2462, y=205)
    pag.typewrite(cno)
    time.sleep(0.5)
    #item name search using notes button
    pag.click(x=2649, y=206)
    time.sleep(0.5)
    #open notes
    pag.click(x=2649, y=206)
    time.sleep(0.5)
    ##notes select all, notes copy
    pag.click(x=3446, y=435)
    pag.keyDown('ctrl')
    pag.keyDown('a')
    pag.keyUp('a')
    pag.keyUp('ctrl')
    time.sleep(0.5)
    
    pag.keyDown('ctrl')
    pag.keyDown('c')
    pag.keyUp('c')
    pag.keyUp('ctrl')
    time.sleep(0.5)
    
    #move to excel workbook || add new sheet
    pag.click(x=-215, y=732)
    pag.keyDown('shift')
    pag.keyDown('f11')
    pag.keyUp('f11')
    pag.keyUp('shift')
    
    ##move to A1 cell
    pag.click(x=-1027, y=216)
    time.sleep(0.5)
    #notes paste
    pag.keyDown('ctrl')
    pag.keyDown('v')
    pag.keyUp('v')
    pag.keyUp('ctrl')
    time.sleep(0.5)
    
    
    #rename sheet
    pag.click(x=-215, y=732)
    pag.keyDown('alt')
    pag.keyDown('h')
    pag.keyUp('h')
    pag.keyDown('o')
    pag.keyUp('o')
    pag.keyDown('r')
    pag.keyUp('r')
    pag.keyUp('alt')
    time.sleep(0.5)
    pag.typewrite(c)
    time.sleep(0.5)
    pag.click(x=-1009, y=219)
    
    #back to notes, close notes
    pag.click(x=3446, y=435)
    time.sleep(0.5)
    pag.click(x=3479, y=162)
    time.sleep(0.5)
    #item nameclear
    pag.click(x=2407, y=179)
    time.sleep(0.5)
    
    
    
    
   