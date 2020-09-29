# -*- coding: utf-8 -*-
"""
Created on Tue Jul 21 13:19:42 2020

@author: SankalpMishra
"""

import numpy as np
import pandas as pd

import xlwings as xw
import time

#%%
start_time = time.time()
path = r'\\shufordyarnsllc.local\SYDFS\UserFiles$\SankalpMishra\Desktop\Sankalp_all\cost_assistanceproject_8_19_2019\Hickory Spinners'

path = path.replace("\\", "/")

col_cno = ['cno', 'yno', 'ply', 'blend', 'tr', 'inv_description', 'case', 'units', 'tare', 'customer']

#df_cno = pd.read_excel(path + 'Cheat Sheet 7-19-20' + '.xlsx', sheet_name = 'ALL YARNS' , usecols = [0,1,2,3,4,5,6,7,8,9])
#df_cno.columns = col_cno
#df_cno['cno'] = df_cno['cno'].astype(str)


#%%
#df_select = df_cno[df_cno['cno'].isin(list(df_input['cno']))]

#%%not used
def proc_pkgwt(inv_desc):
    inv = inv_desc.lower()
    inv = inv.replace('"', '')
    wax_tokens = ['wax', 'wx']
    bag_tokens = ['bag', 'bg']
    cond_tokens = ['cond', 'cd']
    
    l_pkgtype = ['351', '557', '5406rn', 'owt', 'pt', 'pencil tube', 'penciltube', '190', \
                 '190dt', 'cs', '150', '150dt', 'dt', 'OWT5406RN', 'mf', '190mf', '150mf']
    l_inv = inv.split(' ')
    
    for token in l_inv:
        if 'x' in list(token) and 'w' not in list(token) and len(token)> 1: #wax should contain a w with x but dimension token will only have x
            pkg_len = token.split('x')[0]
            pkg_dia = token.split('x')[1]
        
    if token == 'wax' or token == 'wx':
        pass
    
    return l_inv, pkg_len, pkg_dia
#%% Not used
def twist_breakdown(twist):
    try:
        tw_sp = int(''.join(list(str(twist))[:2]))# don't divide by 10. It is done in integrated standards sheet
    except:
        tw_sp = 0
    try:
        tw_tw = int(''.join(list(str(twist))[2:4]))
        if tw_tw >15:
            tw_tw = tw_tw/10 # this division by 10 is not done in the standards sheet
        else:
            pass
    except:
        tw_tw = 0
    try:
        if len(list(twist))>4:
            tw_cable = tw_tw = int(''.join(list(str(twist))[4:6]))
            if tw_tw >15:
                tw_cable = tw_cable/10 # this division by 10 is not done in the standards sheet
            else:
                pass
        else:
            tw_cable = 0
    except:
        tw_cable = 0
        
    return tw_sp, tw_tw, tw_cable

#%%
def putup_breakdown(putup):
    ls_dt = ['190', '150']
    ls_tube = ['owt']
    ls_cone = ['557', '351']
    pw_tube = ['pwt', 'pw']
    
    putup = str(putup).lower()[:3]
    
    if putup in ls_dt:
        return 'dyetube'
    elif putup in ls_tube:
        return 'tube'
    elif putup in ls_cone:
        return 'cone'
    elif putup in pw_tube:
        return 'pwt'
    else:
        return 'default'



def blend_extract(cno):
    try:
        return str(cno)[:3]
    except:
        return 'default'
#%%

df_input = pd.read_excel(path + '/mar_2020/mar20_hs_pp' + '.xlsx', sheet_name= 'hs_pp')
df_input['cno'] = df_input['Item Number'].map(lambda x: str(x)[:-6])
df_input['cno'] = df_input['cno'].map(lambda x: str(x)[-6:])

#%%
df_cno = pd.read_csv(path+ '/mar_2020/dfall_hs_mar20.csv')
df_cno['cno'] = df_cno['cno_itemnum'].map(lambda x: str(x)[-6:])

#%%
wb = xw.Book(r'//shufordyarnsllc.local/SYDFS/UserFiles$/SankalpMishra/Desktop/Sankalp_all/standards_costing_program/HS/tbl/HS_Standards_all.xlsm')
sheet = wb.sheets['Main']

app = wb.app
#blend_summary_mcro = app.macro("carding_drawing_roving_summary") -  example from dudley shoals plant script
summary_macro = app.macro("HS_summary")


#%%
purchased_list = ['ct1', 'ct2', 'pyn']
ls_notfound = []

for item in list(df_input['cno']):
    try:
        print(item)
        df_current = df_cno[df_cno['cno'] == item]
        
        cno_item = df_current['cno_itemnum'].values[0]
        cno_yno = float(str(cno_item)[4:8])/100
        cno_ply = df_current['ply'].values[0]
        tw_sp = df_current['tw_sp'].values[0]
        tw_tw = df_current['tw_tw'].values[0]
        cno_wax = df_current['cno_wax'].values[0]
        cno_bag = df_current['cno_bag'].values[0]
        cno_cond = df_current['cno_cond'].values[0]
        cno_doubspeed = df_current['cno_doubspeed'].values[0]
        cno_twrpm = df_current['cno_twistrpm'].values[0]
        cno_spinrpm = df_current['cno_spinspeed'].values[0]
        cno_pkg = df_current['cno_pkg'].values[0]
        cno_putup = df_current['cno_putup'].values[0]
        cno_putup = putup_breakdown(cno_putup)
        abbrev = blend_extract(cno_item)
        cno_oil = df_current['cno_oil'].values[0]
        
        print('base variables extracted')
        print(cno_item, cno_yno, cno_ply, tw_sp, tw_tw, cno_wax, cno_bag, cno_cond, cno_doubspeed, \
              cno_twrpm, cno_spinrpm, cno_pkg, cno_putup, abbrev, cno_oil)
    #%%Basic setup
        sheet.range("A4").value = cno_item
        sheet.range("B4").value = abbrev
        sheet.range("C4").value = tw_sp
        sheet.range("C5").value = tw_tw
        sheet.range("E4").value = cno_yno
        sheet.range("E5").value = cno_pkg
        sheet.range("G4").value = cno_ply
        
        if str(cno_wax).lower() == 'wx':
            sheet.range("J4").value = 'Yes'
        else:
            sheet.range("J4").value = 'No'
            
        if str(cno_bag).lower() == 'bg':
            sheet.range("K4").value = 'Yes'
        else:
            sheet.range("K4").value = 'No'
            
        pkg_wt = sheet.range("H4").value
        
        
            
        
        if cno_ply > 1:
            creel_wt = pkg_wt/1.8
            # not included creel per crate variable as it is not used in HS doubling/ twisting sheets
            
        else:
            creel_wt = pkg_wt
            
        sheet.range("J6").value = creel_wt # possible error point if a construction number is skipped!
        print('basic sheet setup done')
        #%%carding setup
        
        sheet.range("A9").value = 'Yes'
        sheet.range("B10").value = abbrev
        print('carding setup complete')
        #%% Drawing setup
        
        sheet.range("D9").value = "Yes"
        sheet.range("E10").value = abbrev
        print('drawing setup complete')
        #%%ACO8 Spin setup
        
        sheet.range("G9").value = "Yes"
        #sheet.range("H10").value == ___ #skipping blend match for aco as it doesn't affect output
        
        if cno_ply == 1:
            sheet.range("H14").value = 'Sales'
            sheet.range("H21").value = 'Wood Pallet TP-Sales'
        else:
            sheet.range("H14").value = 'Doubler'
            sheet.range("H21").value = 'Crate'
            
        if cno_putup == 'dyetube':
            sheet.range("H20").value = 'Case-DT'
            sheet.range("H22").value = 'Yes'
        else:
            sheet.range("H20").value = 'Crate-PL' # not much difference between case-pa and crate-pl
            sheet.range("H22").value = 'No'
            
        sheet.range("H15").value = cno_spinrpm
        sheet.range("H19").value = 0.9
        aco_efftemp = sheet.range("H27").value
        
        sheet.range("H19").value = aco_efftemp
        
        print('aco8 setup complete')
        #%% Doubler setup
        if cno_ply>1:
            
            sheet.range("A20").value = 'Yes'
        else:
            sheet.range("A20").value = 'No'
            
        sheet.range("B24").value = cno_doubspeed
        rcom_machine = sheet.range("C27").value
        
        sheet.range("B27").value = rcom_machine
        
        print('doubler setup complete')
        
        #%% Twisting setup
        if cno_ply > 1:
            sheet.range("A37").value = "Yes"
        else:
            sheet.range("A37").value = "No"
            
        # recommended machine --> sheet.range("B42").value = rcom_machine
        sheet.range("B43").value = cno_twrpm
        
        if cno_putup == 'dyetube':
            sheet.range("B46").value = "Yes" # tieoff
            sheet.range("B51").value = 'No' # label
        else:
            sheet.range("B46").value = "No"
            sheet.range("B51").value = 'Yes'
            
        if str(cno_oil).lower() == 'yes':
            sheet.range("B52").value = 'Yes'
        else:
            sheet.range("B52").value = 'No'            
        
        print('twist setup complete')
        #%% Pencil Winder setup
        if cno_putup == 'pwt':
            sheet.range("D20").value = 'Yes'
        else:
            sheet.range("D20").value = 'No'
        
        print('pencil winder setup complete')
        #%%conditioning setup
        if str(cno_cond).lower() == 'cd':
            sheet.range("D44").value = 'Yes'
            sheet.range("E47").value = 2
            
        else:
            sheet.range("D44").value = 'No'
            
        print('condition setup complete')
        #%% shipping/ receiving setup
        sheet.range("G37").value = 'Yes'
            
        print('shipping setup complete')
        #%%Pyn /CT1 removal, Macro run
        if str(abbrev).lower() in purchased_list:
            sheet.range("A9").value = "No"
            sheet.range("D9").value = "No"
            sheet.range("G9").value = "No"
        
        print('running summary macro')
        
        summary_macro()
        print('--------||---------------||---------'*4)
#%%
    except:
        print(item, 'not_found')
        ls_notfound.append(item)
#%%


end_time = time.time()
print('numbers not found:')
print(ls_notfound)
print('script runtime', (end_time - start_time)/60, 'minutes')
    
    
    