# -*- coding: utf-8 -*-
"""
Created on Wed Jul 29 09:30:35 2020

@author: SankalpMishra
"""

import numpy as np
import pandas as pd
import time
#%%
path = r'\\shufordyarnsllc.local\SYDFS\UserFiles$\SankalpMishra\Desktop\Sankalp_all\cost_assistanceproject_8_19_2019\Hickory Spinners'

path = path.replace("\\", "/")

xlsx = pd.ExcelFile(path + '/mar_2020/hs_mar20_setups.xlsx')
cons_sheet = []
for sheet in xlsx.sheet_names:
    cons_sheet.append(xlsx.parse(sheet))
print(cons_sheet)
#%%
def cno_breakdown(cno):
    try:
        all = cno.split('-')
        abbrev = all[0]
        yno = float(all[1].split('/')[0])/100
        ply = int(all[1].split('/')[1])
        cno = all[2]
        return cno, abbrev, yno, ply
    except:
        return ' format error'
    
def pkg_dimensions(package):
    try:
        package = str(package).replace('lbs','')
        package = str(package).replace('lb','')
        package = str(package).replace('oz','')
        package = str(package).replace('yds', '')
        package = str(package).replace('yd', '')
        package = str(package).replace('#', '')
        #all_package = str(package).lower().split(' ')
        dim_list = str(package).lower().split('x')
        if len(dim_list)<3:
            package ='8x' + package
        return package
    except:
        return 'format_error'
    
def twist_breakdown(twist):
    try:
        tw_sp = int(''.join(list(str(twist))[:2]))# don't divide by 10. It is done in integrated standards sheet
    except:
        tw_sp = 0
    try:
        tw_tw = int(''.join(list(str(twist))[-2:]))
        if tw_tw >15:
            tw_tw = tw_tw/10 # this division by 10 is not done in the standards sheet
        else:
            pass
    except:
        tw_tw = 0
    return tw_sp, tw_tw
#%%

list_df = []
ls_df = []
for x in xlsx.sheet_names:
    dfa = pd.read_excel(path + '/mar_2020/hs_mar20_setups.xlsx', sheet_name = x, header = None)
    #dfa.columns = cols # doesn't work since number of columnsvaries from sheet to sheet
    list_df.append(dfa)
    #%%
    try:
        idx_itemnum = dfa.index[dfa[0].str.contains('Item') == True].values[0]
        cno_itemnum = dfa.iloc[idx_itemnum][1]
        
    except:
        cno_itemnum = 'not_found'
    
   #%% 
    try:
        
        idx_dimensions = dfa.index[dfa[0].str.contains('Dimensions') == True].values[0]
        cno_dim = dfa.iloc[idx_dimensions][1]
    except:
        cno_dim = 'not_found'
        
    cno_pkg = pkg_dimensions(cno_dim)
#%%
    try:
        idx_twist = dfa.index[dfa[0].str.contains('Twist') == True].values[0]
        cno_twist = dfa.iloc[idx_twist][1]
        
    except:
        cno_twist = 'not_found'
        
    tw_sp, tw_tw = twist_breakdown(cno_twist)
        #%%
    try:
        idx_wax = dfa.index[dfa[0].str.contains('Wax') == True].values[0]
        cno_wax = dfa.iloc[idx_wax][1]
        
    except:
        cno_wax = 'not_found'
      #%%  
    try:
        idx_bag = dfa.index[dfa[0].str.contains('Bag') == True].values[0]
        cno_bag = dfa.iloc[idx_bag][1]
        
    except:
        cno_bag = 'not_found'
        
    
   #%%     
    try:
        idx_cond = dfa.index[dfa[0].str.contains('Conditioning') == True].values[0]
        cno_cond = dfa.iloc[idx_cond][1]
        
    except:
        cno_cond = 'not_found'
#%%        
    try:    
        cno, abbrev, yno, ply = cno_breakdown(cno_itemnum)
    except:
        cno, abbrev, yno, ply = 'missing', 'missing', 5, 1
        
#%%    
    try:
        if ply > 1:
            idx_doubspeed = dfa.index[dfa[0].str.contains('WINDING') == True].values[0]
            cno_doubspeed = dfa.iloc[idx_doubspeed+1][1]
            
            if (cno_doubspeed == np.NAN) or (cno_doubspeed == ' ') or (cno_doubspeed == np.nan) or (cno_doubspeed == np.NaN):
                cno_doubspeed = dfa.iloc[idx_doubspeed+1][0]
        
            if cno_doubspeed == np.NAN or cno_doubspeed == ' ' or cno_doubspeed == np.nan or cno_doubspeed == np.NaN:
                cno_doubspeed = dfa.iloc[idx_doubspeed][1]
        else:
            cno_doubspeed = 0   
             
    except:
        cno_doubspeed = 'not_found'
        
    
    
#%%
    if ply>1:
        try:
            idx_twistrpm = dfa.index[dfa[7].str.contains('SPINDLE') == True].values[0]
            cno_twistrpm = dfa.iloc[idx_twistrpm][8]
         
        except:
            try:
                idx_twistrpm = dfa.index[dfa[8].str.contains('SPINDLE') == True].values[0]
                cno_twistrpm = dfa.iloc[idx_twistrpm][9]
                
            except:
                try:
                    idx_twistrpm = dfa.index[dfa[9].str.contains('SPINDLE') == True].values[0]
                    cno_twistrpm = dfa.iloc[idx_twistrpm][10]
                  
                except:
                    cno_twistrpm = 'not_found'
            
    else:
        cno_twistrpm = 0
        
    if cno_twistrpm == 'not_found':
        try:
            idx_twistrpm = dfa.index[dfa[7].str.contains('RPM') == True].values[0]
            cno_twistrpm = dfa.iloc[idx_twistrpm][8]
        except:
            try:
                idx_twistrpm = dfa.index[dfa[8].str.contains('RPM') == True].values[0]
                cno_twistrpm = dfa.iloc[idx_twistrpm][9]
            except:
                try:
                    idx_twistrpm = dfa.index[dfa[9].str.contains('RPM') == True].values[0]
                    cno_twistrpm = dfa.iloc[idx_twistrpm][10]
                except:
                    cno_twistrpm = 'not_found'
                    
    #%% for oilers
    if ply>1:
        try:
            
            idx_oil = dfa.index[dfa[7].str.contains('OILERS') == True].values[0]
            cno_oil = dfa.iloc[idx_oil][8]
        except:
            try:
                idx_oil = dfa.index[dfa[8].str.contains('OILERS') == True].values[0]
                cno_oil = dfa.iloc[idx_oil][9]
            except:
                try:
                    idx_oil = dfa.index[dfa[9].str.contains('OILERS') == True].values[0]
                    cno_oil = dfa.iloc[idx_oil][10]
                except:
                    cno_oil = 'not_found'
    else:
        cno_oil = 'No'
    
    #%%
    try:
        idx_spinspeed = dfa.index[dfa[0].str.contains('Rotor') == True].values[0]
        cno_spinspeed = dfa.iloc[idx_spinspeed][1]
        
    except:
        try:
            idx_spinspeed = dfa.index[dfa[2].str.contains('ROTOR RPM') == True].values[0]
            cno_spinspeed = dfa.iloc[idx_spinspeed][3]
        except:
            
            cno_spinspeed = 'not_found'
            
    #%%
    try: 
        idx_putup = dfa.index[dfa[0].str.contains('Put') == True].values[0]
        cno_putup = dfa.iloc[idx_putup][1]
    except:
        cno_putup = 'not_found'
    #%%
    
    #time.sleep(5)   
    print(cno_itemnum, ply, cno_dim, tw_sp, tw_tw, cno_wax, cno_bag, cno_cond, cno_doubspeed, cno_twistrpm, cno_spinspeed, cno_pkg, cno_putup, cno_oil)
    #print(dfa)
    
    cols = ['cno_itemnum', 'ply', 'cno_dim', 'tw_sp', 'tw_tw', 'cno_wax', 'cno_bag', 'cno_cond', 'cno_doubspeed', 'cno_twistrpm', 'cno_spinspeed', 'cno_pkg', 'cno_putup', 'cno_oil']
    
    dfb = pd.DataFrame([[cno_itemnum, ply, cno_dim, tw_sp, tw_tw, cno_wax, cno_bag, cno_cond, cno_doubspeed, cno_twistrpm, cno_spinspeed, cno_pkg, cno_putup, cno_oil]], columns = cols)
    print(dfb)
    ls_df.append(dfb)
    
df_all = pd.concat(ls_df)
#%%
df_all.to_csv (path + '/mar_2020/dfall_hs_mar20.csv', index = None, header=True) 

