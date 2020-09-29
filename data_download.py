# -*- coding: utf-8 -*-
"""
Created on Thu May 21 14:58:29 2020

@author: SankalpMishra
"""

import numpy as np
import pandas as pd
#import openpyxl  xl wings is better for real time excel processing
import xlwings as xw
import re
from fuzzywuzzy import fuzz, process

import time

#%%
starttime = time.time()
path = '//shufordyarnsllc.local/SYDFS/UserFiles$/SankalpMishra/Desktop/Sankalp_all/cost_assistanceproject_8_19_2019/'
path_cno = '//shufordyarnsllc.local/SYDFS/UserFiles$/SankalpMishra/Desktop/Sankalp_all/standards_costing_program/tbl/'

columns_rs = ['cno_complete', 'blend_detail', 'twist', 'package_detail', 'wax', 'bag', 'condition', 'traypack', 'customer']
df_cno_rs = pd.read_excel(path_cno + 'Construction Numbers 2 Dudley.xlsx', sheet_name = 'Ring', header = None, skiprows = range(3),\
                          usecols = [1,2,4,5,6,7,8,9,11])

df_cno_rs.columns = columns_rs

print(df_cno_rs.head())
#%%

columns_oe = ['cno_complete', 'blend_detail', 'twist', 'package_detail', 'wax', 'bag', 'condition', 'traypack', 'customer']
df_cno_oe = pd.read_excel(path_cno + 'Construction Numbers 2 Dudley.xlsx', sheet_name = 'OE', header = None, skiprows = range(3),\
                          usecols = [1,2,3,4,5,6,7,8,9])

df_cno_oe.columns = columns_oe

print(df_cno_oe.head())

#%%
columns_aj = ['cno_complete', 'blend_detail', 'twist', 'package_detail', 'wax', 'bag', 'condition', 'traypack', 'customer']

df_cno_aj = pd.read_excel(path_cno + 'Construction Numbers 2 Dudley.xlsx', sheet_name = 'Air-Jet', header = None,\
                          usecols = [1,2,4,5,6,7,8,9,11])

df_cno_aj.columns = columns_aj

print(df_cno_aj.head())

#%%
df_yno_range_spin = pd.read_excel(path_cno + 'runspeed_tbl.xlsx', sheet_name = 'spin')
df_yno_range_twist = pd.read_excel(path_cno + 'runspeed_tbl.xlsx', sheet_name = 'doub_twist')
df_blend_tier = pd.read_excel(path_cno + 'runspeed_tbl.xlsx', sheet_name = 'blend_tier')
#%%
purchased_list = ['ct1', 'ct2', 'pyn']
dyetube_list = ['190mf', 'mf', 'f', '150', '150mf'] #check for all dt types
cone_list = ['557', '351']
tube_list = ['190mf', 'mf', 'f', '150', '54', '54mmrn', 'pwt',  'owt5406rn', 'owt']


#%% String Split functions

score_cutoff = 80 # cutoff for blend matching in partial match function
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
    

def blend_breakdown(blend):
    remove_lex = ['rs', '19d', 'wtp', 'tp', 'tpp', 'reverse', '90s', '1.5d', 'oe', '2.0d', '1.5d/2.0d', 'aj', 'nanya', 'polyester', 'nanay', 'pakistan']
    blend = str(blend).replace('-', '')
    print_blend = []
    #blend = blend.replace('/', ' ')# using token sort ratio doesn't account for "/" anyway
    try:
        blend_fil = re.sub(r'([^a-zA-Z\s]+?)','', blend)
        blend_fil = blend_fil.lower().split(' ')
        blend_fil = [bl for bl in blend_fil if bl != '']
        blend_fil.sort()
        
        all_lex = blend.strip().split(' ')[1:]
        ls = []
        for lex in all_lex:
            if (lex.strip().lower() in remove_lex):
                pass
            else:
                ls.append(lex)
                ls.sort()
                print_blend.append(lex)
         
            
        return ' '.join(ls), ls, blend_fil, ''.join(print_blend)
    
    except:
        return 'format error'
    
def blend_breakdown_fordataset(blend):
    remove_lex = ['rs', '19d', 'wtp', 'tp', 'tpp', 'reverse', '90s', '1.5d', 'oe', '2.0d', '1.5d/2.0d', 'aj', 'nanya','polyester', 'nanay']
    blend = str(blend).replace('-', '')
    #blend = blend.replace('/', ' ')
   
    try:
        blend_fil = re.sub(r'([^a-zA-Z\s]+?)','', blend)
        blend_fil = blend_fil.lower().split(' ')
        blend_fil = [bl for bl in blend_fil if bl != '']
        blend_fil.sort()
        
        all_lex = blend.strip().split(' ')
        ls = []
        for lex in all_lex:
            if (lex.strip().lower() in remove_lex):
                pass
            else:
                ls.append(lex)
                ls.sort()
         
            
        return ' '.join(ls), ls, blend_fil
    
    except:
        return 'format error'
            
   
def package_breakdown(package):
    
    try:
        package = str(package).replace('#','')
        package = str(package).replace('oz','')
        all_package = str(package).lower().split(' ')
        
            
        
        print(all_package)
        package_type = all_package[0]
        
        put_up = str(all_package[1]).replace('pw', '')
        package_dim = all_package[1].split('x')
        
        package_wt = package_dim[2]
        package_wt = float(re.sub(r'([^\d\.])','', package_wt))
        if package_wt > 15:
            package_wt = package_wt = package_wt/16
        
        
        return package_type, float(package_dim[0]), float(package_dim[1]), package_wt, put_up
    
    except:
        return 'format error'


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

def partial_match(blend, choices):
    one = process.extractOne(blend, choices, scorer = fuzz.token_sort_ratio, score_cutoff = score_cutoff)
    if one != None:
        return process.extractOne(blend, choices, scorer = fuzz.token_sort_ratio, score_cutoff = score_cutoff)
    else:
        return ("Default", 0)
    
        
        
#%% getting blend summaries for pre blend from summary sheet

df_preblend_data = pd.read_excel(path_cno + 'DS_MAIN_STANDARDS_new_by_department - blend_macro_test_recovered.xlsm', \
                                 sheet_name = 'Data on Blends', skiprows = range(2), usecols = [1,2,3,4,5])

preblenddata_cols = ['blend_pre', 'abbrev_pre', 'balewt_pre','rt_pre', 'explb_pre']

df_preblend_data.columns = preblenddata_cols

df_preblend_data['blend_pre_list'] = df_preblend_data['blend_pre'].map(lambda x: blend_breakdown_fordataset(x)[0])

preblend_choices = list(df_preblend_data['blend_pre_list'])

preblend_abbrev_list = ['ptx', 'pow', 'ace', 'ps8', 'a82', 'ptr']
#%% getting carding blend data from carding summary sheet
df_carding_data = pd.read_excel(path_cno + 'DS_MAIN_STANDARDS_new_by_department - blend_macro_test_recovered.xlsm', \
                                sheet_name = 'blend_input_cards')
df_carding_data['blend'] = df_carding_data['Full Description'].map(lambda x: blend_breakdown_fordataset(x)[0])

carding_choices = list(df_carding_data['blend'])

#%%getting drawing data fromd rawing summary sheet
df_drawing_data = pd.read_excel(path_cno + 'DS_MAIN_STANDARDS_new_by_department - blend_macro_test_recovered.xlsm', \
                                sheet_name = 'blend_input', use_cols= [0,1,2,3,4,5], header = 2)
df_drawing_data['blend'] = df_drawing_data['Description'].map(lambda x: blend_breakdown_fordataset(x)[0])

drawing_choices = list(df_drawing_data['blend'])

#%% getting roving data from roving summary sheet
df_roving_data = pd.read_excel(path_cno + 'DS_MAIN_STANDARDS_new_by_department - blend_macro_test_recovered.xlsm', \
                                sheet_name = 'blend_input', use_cols= [0,1,2,3,4,5], header = 2)
df_roving_data['blend'] = df_roving_data['Description'].map(lambda x: blend_breakdown_fordataset(x)[0])

roving_choices = list(df_roving_data['blend'])

#%% loadin the monthly production pounds data
df_input = pd.read_excel(path+ '/Aug_2020/' + 'Aug20_pp.xlsx', sheet_name = 'aug_pp') #for left over numbers
df_input['cno'] = df_input['Item Number'].map(lambda x: str(x)[:-6])

#%% generating input variables from input dataframe
ls_notfound = []
ls_found = []
inputdf_columns = ['cno', 'abbrev','blend', 'yno', 'ply', 'tw_sp', 'tw_tw'\
                                     , 'package_type', 'package_dim', 'wax', 'bag', 'condition', \
                                     'traypack', 'customer']

wb = xw.Book(r'//shufordyarnsllc.local/SYDFS/UserFiles$/SankalpMishra/Desktop/Sankalp_all/standards_costing_program/tbl/DS_MAIN_STANDARDS_new_by_department - blend_macro_test_recovered.xlsm')
sheet = wb.sheets['Main']

app = wb.app
blend_summary_mcro = app.macro("carding_drawing_roving_summary")
summary_macro = app.macro("MainSummary")

#blend_summary_mcro()
#%%


for i, row in df_input.iterrows():
#for i in range(1):
    cno_all = df_input.iloc[i]['cno']
    print(cno_all)
    #for custom cno setting:
    #cno_all = 'PTX-2050/03-090401'
    if cno_all in list(df_cno_rs['cno_complete']):
        blend_detail = df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'blend_detail'].values[0]
        blend = blend_breakdown(blend_detail)[0]
        blend_print = blend_breakdown(blend_detail)[3]
        tw_sp = twist_breakdown(df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'twist'].values[0])[0]
        tw_tw = twist_breakdown(df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'twist'].values[0])[1]
        package_dim = package_breakdown(df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'package_detail'].values[0])[4]
        package_weight = package_breakdown(df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'package_detail'].values[0])[3]
        package_length = package_breakdown(df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'package_detail'].values[0])[1]
        package_type = package_breakdown(df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'package_detail'].values[0])[0]
        
        wx = df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'wax'].values[0]
        bg = df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'bag'].values[0]
        condition = df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'condition'].values[0]
        traypack = df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'traypack'].values[0]
        customer = df_cno_rs.loc[df_cno_rs['cno_complete']== cno_all, 'customer'].values[0]
        
        cno_num, abbrev, yno, ply = cno_breakdown(cno_all)
        ls_found.append([cno_all, abbrev, blend, yno, ply, tw_sp, tw_tw, package_type, package_dim, wx, bg,\
                                   condition, traypack, customer])
        df_current = pd.DataFrame([[cno_all, abbrev, blend, yno, ply, tw_sp, tw_tw, package_type, package_dim, wx, bg,\
                                   condition, traypack, customer]], columns = inputdf_columns)
        spin = 'rs'
        passes = 2
        preblend_match = partial_match(blend, preblend_choices)
        carding_match = partial_match(blend, carding_choices)
        drawing_match = partial_match(blend, drawing_choices)
        roving_match = partial_match(blend, roving_choices)
        print(df_current, preblend_match, carding_match, drawing_match, roving_match)
        
        
    elif cno_all in list(df_cno_oe['cno_complete']):
        blend_detail = df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'blend_detail'].values[0]
        blend = blend_breakdown(blend_detail)[0]
        blend_print = blend_breakdown(blend_detail)[3]
        tw_sp = twist_breakdown(df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'twist'].values[0])[0]
        tw_tw = twist_breakdown(df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'twist'].values[0])[1]
        package_dim = package_breakdown(df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'package_detail'].values[0])[4]
        package_weight = package_breakdown(df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'package_detail'].values[0])[3]
        package_length = package_breakdown(df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'package_detail'].values[0])[1]
        package_type = package_breakdown(df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'package_detail'].values[0])[0]
        
        wx = df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'wax'].values[0]
        bg = df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'bag'].values[0]
        condition = df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'condition'].values[0]
        traypack = df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'traypack'].values[0]
        customer = df_cno_oe.loc[df_cno_oe['cno_complete']== cno_all, 'customer'].values[0]
        
        cno_num, abbrev, yno, ply = cno_breakdown(cno_all)
        ls_found.append([cno_all, abbrev, blend, yno, ply, tw_sp, tw_tw, package_type, package_dim, wx, bg,\
                                   condition, traypack, customer])
    
        df_current = pd.DataFrame([[cno_all, abbrev, blend, yno, ply, tw_sp, tw_tw, package_type, package_dim, wx, bg,\
                                   condition, traypack, customer]], columns = inputdf_columns)
        spin = 'oe'
        passes = 1
        preblend_match = partial_match(blend, preblend_choices)
        carding_match = partial_match(blend, carding_choices)
        drawing_match = partial_match(blend, drawing_choices)
        print(df_current, preblend_match, carding_match, drawing_match)
    
    elif cno_all in list(df_cno_aj['cno_complete']):
        blend_detail = df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'blend_detail'].values[0]
        blend = blend_breakdown(blend_detail)[0]
        blend_print = blend_breakdown(blend_detail)[3]
        tw_sp = twist_breakdown(df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'twist'].values[0])[0]
        tw_tw = twist_breakdown(df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'twist'].values[0])[1]
        package_dim = package_breakdown(df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'package_detail'].values[0])[4]
        package_weight = package_breakdown(df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'package_detail'].values[0])[3]
        package_length = package_breakdown(df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'package_detail'].values[0])[1]
        package_type = package_breakdown(df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'package_detail'].values[0])[0]
        
        wx = df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'wax'].values[0]
        bg = df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'bag'].values[0]
        
        
        condition = df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'condition'].values[0]
        traypack = df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'traypack'].values[0]
        customer = df_cno_aj.loc[df_cno_aj['cno_complete']== cno_all, 'customer'].values[0]
        
        cno_num, abbrev, yno, ply = cno_breakdown(cno_all)
        ls_found.append([cno_all, abbrev, blend, yno, ply, tw_sp, tw_tw, package_type, package_dim, wx, bg,\
                                   condition, traypack, customer])
    
        df_current = pd.DataFrame([[cno_all, abbrev, blend, yno, ply, tw_sp, tw_tw, package_type, package_dim, wx, bg,\
                                   condition, traypack, customer]], columns = inputdf_columns)
        spin = 'aj'
        preblend_match = partial_match(blend, preblend_choices)
        carding_match = partial_match(blend, carding_choices)
        drawing_match = partial_match(blend, drawing_choices)
        passes = 3
        print(df_current, preblend_match, carding_match, drawing_match)
    
    else:
        ls_notfound.append(cno_all)
        spin = 'not_found'
        
    ls_found.append([cno_all, abbrev, blend, yno, ply, tw_sp, tw_tw, package_type, package_dim, wx, bg,\
                                   condition, traypack, customer])   
    #opening the engine workbook for standards
#    engine_book = openpyxl.load_workbook(path + '/tbl/' +  'DS_MAIN_STANDARDS_new_by_department.xlsx')
#    main_sheet = engine_book["Main"]
    
    ##main_sheet["Z1"] = "Test"
    ##engine_book.save(path + '/tbl/' + '/bin/'+  'DS_MAIN_STANDARDS_new_by_department_test.xlsx')
    ######  #####    
    
    print("main variables setup")
    if spin != 'not_found':
        
        if (str(wx).lower() == 'wx'): wax = 'Yes'
        else: wax = 'No'
        
        if (str(bg).lower() == 'wx'): bag = 'Yes'
        else: bag = 'No'
        
        yno_net = yno/ply
        
        df_speed_spin = df_yno_range_spin[(yno > df_yno_range_spin['yno_l']) & (yno <= df_yno_range_spin['yno_u'])]
        df_speed_twist = df_yno_range_twist[(yno_net > df_yno_range_twist['yno_l']) & (yno_net <= df_yno_range_twist['yno_u'])]
        
        sheet.range("A5").value = cno_all
        sheet.range("B5").value = blend
        sheet.range("C5").value = tw_sp
        sheet.range("C6").value = tw_tw
        sheet.range("E5").value = yno
        sheet.range("E6").value = package_dim
        sheet.range("G5").value = ply
        sheet.range("J5").value = wax
        sheet.range("K5").value = bag
        sheet.range("B35").value = 0.9 #aco8 estimated efficiency, should be some value to get 0 when not selecting Aco8, else shows DIV!0 error
        sheet.range("E35").value = 0.9 # above statement for airjet
        
        package_per_crate = sheet.range("I71").value
        if ply > 1:
            creel_weight = package_weight/1.8
            creel_per_crate = package_per_crate - 32
        else:
            creel_weight= package_weight
            creel_per_crate = package_per_crate
            
        

#%% Preblend variables setup
        print("Preblend setup")
        if str(abbrev).lower() in preblend_abbrev_list:
            sheet.range("A8").value = "Yes"
            sheet.range("B9").value = df_preblend_data.loc[df_preblend_data['blend_pre_list'] == preblend_match[0], 'blend_pre'].values[0]
        else: 
            sheet.range("A8").value = "No"
            
#%% Carding variables setup
        print("Carding setup")
        sheet.range("A14").value = "Yes"
        sheet.range("B15").value = df_carding_data.loc[df_carding_data['blend'] == carding_match[0], 'Full Description'].values[0]
        blend_tier = df_carding_data.loc[df_carding_data['blend'] == carding_match[0], 'tier'].values[0]
        
#%% Drawing variable setup
        print("Drawing setup")
        sheet.range("D8").value = "Yes"
        sheet.range("E9").value = df_drawing_data.loc[df_drawing_data['blend'] == drawing_match[0], 'Description'].values[0]
        sheet.range("E10").value = passes
        
#%% Roving variables setup
        print("Roving setup")
        if spin == 'rs':
            sheet.range("G8").value = "Yes"
            sheet.range("H9").value = df_roving_data.loc[df_roving_data['blend'] == roving_match[0], 'Description'].values[0]
            
        else:
            sheet.range("G8").value = "No"
#%% OE spin setup. Aco8 for yno>12, else SE9
        print("OE spin setup")
        if spin == 'oe' and yno > 10:
            
            speed_aco = df_speed_spin['aco'].iloc[0] + df_blend_tier.loc[df_blend_tier['tier'] == blend_tier, 'oe'].values[0]
            sheet.range("A25").value = "Yes"
            sheet.range("M25").value = "No"
            
            sheet.range("B31").value = speed_aco
            if ply >1:
                sheet.range("B30").value = "Doubler"
                sheet.range("B37").value = "Crate"
                
            else:
                sheet.range("B30").value = "Sales"
                sheet.range("B37").value = "Wood Pallet TP-Sales"
                
            if package_type in dyetube_list:
                sheet.range("B36").value = "Case-DT"
            else:
                sheet.range("B36").value = "Crate-PL"
            
            if package_type in tube_list:
                sheet.range("B38").value = "Yes"
            else:
                sheet.range("B38").value = "No"
            
            sheet.range("B35").value = 0.9
            #time.sleep(1)
            aco_efftemp = sheet.range("B43").value
            #time.sleep(1)  #sleep added because excel is not updating values quick enough resulting in errors
            sheet.range("B35").value = aco_efftemp
            #time.sleep(2)
                
        elif spin == 'oe' and yno <= 12: # se9 variable setup
            
            speed_se9 = df_speed_spin['se9'].iloc[0]
            
            sheet.range("A25").value = "No"
            sheet.range("M25").value = "Yes"
            
            
            sheet.range("N30").value = speed_se9
        else:
            sheet.range("A25").value = "No"
            sheet.range("M25").value = "No"
            
#%%  Airjet variables setup
        print("Air-jet setup")
            
        if spin == 'aj':
            
            sheet.range("D25").value = "Yes"
            
            speed_aj = df_speed_spin['aj'].iloc[0]
            
            sheet.range("E29").value = speed_aj
            
#           sheet.range("E26").value = # blend match not very significant in current mjs standards. :/
            if package_type in dyetube_list:
                sheet.range("E30").value = "Case-DT"
            else:
                sheet.range("E30").value = "Case-PT"
                
            if ply>1:
                sheet.range("E32").value = 'Wood Tray Pack'
            else:
                sheet.range("E32").value = 'Crate'
                
            if package_type in tube_list:
                sheet.range("E34").value = "Yes"
            else:
                sheet.range("E34").value = "No"
            
            sheet.range("E35").value = 0.9
            #time.sleep(1)
            aj_efftemp = sheet.range("E41").value
            #time.sleep(1)  #sleep added because excel is not updating values quick enough resulting in errors
            sheet.range("E35").value = aj_efftemp
                
            
        else:
            sheet.range("D25").value = "No"
            
        
#%% Ring spinning variables setup: Marzoli and zinzer in order
        print("Ring spin setup")
        if spin == 'rs' and yno>16:
            
            speed_marzoli_spin = df_speed_spin['marzoli_spin'].iloc[0] + df_blend_tier.loc[df_blend_tier['tier'] == blend_tier, 'marzoli_spin'].values[0]
            speed_marzoli_wind = df_speed_spin['marzoli_wind'].iloc[0] + df_blend_tier.loc[df_blend_tier['tier'] == blend_tier, 'marzoli_wind'].values[0]
            
            sheet.range("G25").value = 'Yes'
            sheet.range("J25").value = 'No' #No in strict yno boundary conditions on wo machine
#siro            sheet.range("H29")
            sheet.range("H30").value = speed_marzoli_spin
            
            sheet.range("H33").value = speed_marzoli_wind

        elif spin == 'rs' and yno <= 16: # zinser variable setup
            
            speed_zinser_spin = df_speed_spin['zinser_spin'].iloc[0]+ df_blend_tier.loc[df_blend_tier['tier'] == blend_tier, 'zinser_spin'].values[0]
            speed_zinser_wind = df_speed_spin['zinser_wind'].iloc[0]+ df_blend_tier.loc[df_blend_tier['tier'] == blend_tier, 'zinser_wind'].values[0]
            
            sheet.range("G25").value = 'No' #No in strict yno boundary conditions on wo machine
            sheet.range("J25").value = 'Yes'
            
            if package_type in tube_list:
                sheet.range("K30").value = 'Dye Tube'
            else:
                sheet.range("K30").value = 'Cone'
            
            sheet.range("K29").value = speed_zinser_spin
            
            sheet.range("K32").value = speed_zinser_wind
            
            
            
        else:
            sheet.range("G25").value = 'No'
            sheet.range("J25").value = 'No'
            
#%% Doubling variable setup
        print("Doubling setup")
        if ply > 2:
            
            speed_doub = df_speed_twist['doubler'].iloc[0]
            
            sheet.range("A49").value = 'Yes'
            sheet.range("B53").value = speed_doub
            
        else:
            sheet.range("A49").value = 'No'
#%% Twisting
        print("Twisting setup")
        if ply>1:
            sheet.range("J49").value = 'Yes'
            sheet.range("M51").value = 1
            
            speed_twist_5 = df_speed_twist['twisting_5'].iloc[0]
            speed_twist_6 = df_speed_twist['twisting_6'].iloc[0]
            speed_twist_7 = df_speed_twist['twisting_7'].iloc[0]
            
            if 0 < package_length <6.5:
                sheet.range("K54").value = 7
                sheet.range("K56").value = '6 in. tube'
                sheet.range("K55").value = speed_twist_7
                
            elif 6.5 <= package_length < 8.5:
                sheet.range("K54").value = 6
                sheet.range("K56").value = '7 in. tube'
                sheet.range("K55").value = speed_twist_6
                
            elif package_length >= 8.5:
                sheet.range("K54").value = '05+7'
                sheet.range("K56").value = '10 in. tube'
                sheet.range("K55").value = speed_twist_5
                
            else:
                sheet.range("K54").value = 7
                sheet.range("K56").value = '6 in. tube'
                sheet.range("K55").value = speed_twist_7
            
            #doff package to: K57: set to wood tray pack always after twisting unless observed otherwise
            
            if ply == 2:
                sheet.range("M51").value = 2
                
            else:
                sheet.range("M51").value = 1
                
            sheet.range("M52").value  = creel_per_crate
            
            if package_type in cone_list:
                if (0 < package_length<=6.5): sheet.range("M53").value = '6 in. cone'
                elif (6.5 < package_length < 8) : sheet.range("M53").value = '8 in. cone'
                elif (8 <= package_length < 10.5) : sheet.range("M53").value = '10 in. cone'
            elif package_type in dyetube_list:
                sheet.range("M53").value = '6 in. DT190NS'
            else:
                if (package_length<8): sheet.range("M53").value = '10 in. paper tube'
                else: sheet.range("M53").value = '6 in. cone'
            
            sheet.range("M59").value = 0.9
            twist_efftemp = sheet.range("K61").value
            #time.sleep(1)  #sleep added because excel is not updating values quick enough resulting in errors, not really
            sheet.range("M59").value = twist_efftemp
            
        else:
            sheet.range("J49").value = 'No'
            
#%% back winding and Xeno variable setup, duro and conditioning setup
        print("Xeno, Duro & Conditioning setup")
        # backwindign check??
        
        ##xeno setup
        
        if 'p' in list(package_type) and 'w' in package_type and 't' in package_type:
            
            speed_xeno = df_speed_twist['xeno'].iloc[0]
            sheet.range("A67").value = 'Yes'
            sheet.range("B71").value = speed_xeno         
            # numspin =5
        else:
            sheet.range("A67").value = 'No'           
            
        
        if package_type == '330':
            speed_duro = df_speed_twist['duro'].iloc[0]
            sheet.range("D67").value = 'Yes'
            sheet.range("E71").value = speed_duro
            
        else:
            sheet.range("D67").value = 'No'
            
        if condition == 'cd':
            sheet.range("G67").value = 'Yes'
            if str(customer).lower() in ['niedner', 'neidner']:
                sheet.range("H70").value = 3
            else:
                sheet.range("H70").value = 2
        else:
            sheet.range("G67").value = 'No'
            
        print("setting purchased-yarn spin values to 0")
            
        if str(abbrev).lower() in purchased_list:
            sheet.range("A8").value = "No"
            sheet.range("A14").value = "No"
            sheet.range("D8").value = "No"
            sheet.range("G8").value = "No"
            sheet.range("A25").value = "No"
            sheet.range("D25").value = "No"
            sheet.range("J25").value = "No"
            sheet.range("G25").value = "No"
            sheet.range("M25").value = "No"
        print("Running summary macro")
        summary_macro()           
    
        print("--------||--------"*5)  
#%%
              
    
    #preb = 
        
        shipping_explbs = sheet['K74'].value
#    wb.close()
    
    
#    df_a = pd.read_excel(path + '/tbl/' + 'DS_MAIN_STANDARDS_new_by_department.xlsx', sheet_name = 'Main', header = None)
#    print(df_a)
#    shipping_explbs = df_a.iloc[73][10]
        print(shipping_explbs)


    else: 
        pass
    
df_allinput = pd.DataFrame(ls_found, columns = inputdf_columns)
#engine_book.save(path + '/tbl/' +  'DS_MAIN_STANDARDS_new_by_department.xlsx')
#engine_book.close()
#%%

def find_cno(cno):
    if not df_cno_oe[df_cno_oe['cno_complete']== cno].empty:
        return df_cno_oe[df_cno_oe['cno_complete']== cno], 'oe'
    
    elif not df_cno_rs[df_cno_rs['cno_complete']== cno].empty:
        return df_cno_rs[df_cno_rs['cno_complete']== cno], 'rs'
    
    elif not df_cno_aj[df_cno_aj['cno_complete']== cno].empty:
        return df_cno_aj[df_cno_aj['cno_complete']== cno], 'aj'
    else:
        return 'not found'
    
#wb.close()
print(ls_notfound)      
endtime = time.time()

print((endtime-starttime)/60, 'mins') 
#%%
blend_cno_rs = df_cno_rs['blend_detail'].map(lambda x: blend_breakdown(x)[0])
c = set(blend_cno_rs)
c = list(c)        
        
    



























