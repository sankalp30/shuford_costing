# -*- coding: utf-8 -*-
"""
Created on Tue Sep  1 14:00:49 2020

@author: SankalpMishra
"""

import numpy as np
import pandas as pd
#%%
path = '//shufordyarnsllc.local/SYDFS/UserFiles$/SankalpMishra/Desktop/Sankalp_all/cost_assistanceproject_8_19_2019/Hickory Spinners/'


std_file = 'mar_2020/hs_mar20_stds'
df_standards = pd.read_excel(path+ std_file + '.xlsx', sheet_name = 'Sheet1')
df_standards['cno'] = df_standards['Cno']

factor_file = 'input_tbl/current_utility_factor_HS'
df_factors = pd.read_excel(path+ factor_file + '.xlsx', sheet_name = 'Updated_dep_factors')

trx_file = 'mar_2020/mar20_hs_pp'
df_trx = pd.read_excel(path + trx_file + '.xlsx', sheet_name = 'hs_pp')
df_trx['cno'] = df_trx['Item Number'].map(lambda x: str(x)[:-6])
df_trx['trx_quan'] = df_trx['TRX QTY']

df_cnodet = df_standards.merge(df_trx, on = ['cno'])
#%%

df_cnodet = df_cnodet.fillna(0)

#%%
class cno_det:
    def __init__(self, current_cno, df_currentcno = [], df_factors = []):
        
        self.cotton_list = ['CAT', 'SKY', 'COT', 'CSU', 'CCA']
        self.kwh_rate = (0.053489743/3600)
        self.current_cno = current_cno
        self.df_currentcno = df_currentcno
        
        self.cno = df_currentcno['cno'].values[0]
        self.yno = df_currentcno['Yno'].values[0]
        self.ply = df_currentcno['Ply'].values[0]
        self.pkg_wt = df_currentcno['pkg_wt'].values[0]
        self.trx_quan = df_currentcno['trx_quan'].values[0]
        self.abbrev = df_currentcno['Blend'].values[0]
        
        self.card_explbs = df_currentcno['Carding'].values[0]
        self.draw_explbs = df_currentcno['Drawing'].values[0]
        self.aco_explbs = df_currentcno['Aco_spinning'].values[0]
        self.doub_explbs = df_currentcno['Doubling'].values[0]
        self.twist_explbs = df_currentcno['Twisting'].values[0]
        self.pencilwind_explbs = df_currentcno['Pencil_wind'].values[0]
        self.cond_explbs = df_currentcno['Conditioning'].values[0]
        self.backwind_explbs = df_currentcno['backwind'].values[0]
        self.shipping_explbs = df_currentcno['Shipping'].values[0]
        
        self.cotcard_erate = df_factors[df_factors['dep'] == 'Cotton carding']['elect_persec'].values[0] #combine cotton and pC carding for final costing
        self.polycard_erate = df_factors[df_factors['dep']== 'PC carding']['elect_persec'].values[0]
        self.draw_erate = df_factors[df_factors['dep'] == 'Drawing']['elect_persec'].values[0]
        self.aco_erate = df_factors[df_factors['dep'] == 'OE spinning']['elect_persec'].values[0]
        self.doubsix_erate = df_factors[df_factors['dep'] == '6" doubler']['elect_persec'].values[0]
        self.doubeight_erate = df_factors[df_factors['dep'] == '8" doubler']['elect_persec'].values[0]
        self.doubten_erate = df_factors[df_factors['dep'] == '10" doubler']['elect_persec'].values[0]
        self.twistsix_erate = df_factors[df_factors['dep'] == '6" twister']['elect_persec'].values[0]
        self.twisteight_erate = df_factors[df_factors['dep'] == '8" twister']['elect_persec'].values[0]
        self.twistten_erate = df_factors[df_factors['dep'] == '10" twister']['elect_persec'].values[0]
        
        self.pwt_erate = df_factors[df_factors['dep'] == 'pencil winder']['elect_persec'].values[0]
        self.cond_erate = df_factors[df_factors['dep'] == 'conditioning']['elect_persec'].values[0]
        self.pack_erate = df_factors[df_factors['dep'] == 'packing']['elect_persec'].values[0]
        
        self.cotcard_watrate = df_factors[df_factors['dep'] == 'Cotton carding']['water_persec'].values[0]
        self.polycard_watrate = df_factors[df_factors['dep'] == 'PC carding']['water_persec'].values[0]
        self.draw_watrate = df_factors[df_factors['dep'] == 'Drawing']['water_persec'].values[0]
        self.aco_watrate = df_factors[df_factors['dep'] == 'OE spinning']['water_persec'].values[0]
        self.doubsix_watrate = df_factors[df_factors['dep'] == '6" doubler']['water_persec'].values[0]
        self.doubeight_watrate = df_factors[df_factors['dep'] == '8" doubler']['water_persec'].values[0]
        self.doubten_watrate = df_factors[df_factors['dep'] == '10" doubler']['water_persec'].values[0]
        self.twistsix_watrate = df_factors[df_factors['dep'] == '6" twister']['water_persec'].values[0]
        self.twisteight_watrate = df_factors[df_factors['dep'] == '8" twister']['water_persec'].values[0]
        self.twistten_watrate = df_factors[df_factors['dep'] == '10" twister']['water_persec'].values[0]
        
        self.pwt_watrate = df_factors[df_factors['dep'] == 'pencil winder']['water_persec'].values[0]
        self.cond_watrate = df_factors[df_factors['dep'] == 'conditioning']['water_persec'].values[0]
        self.pack_watrate = df_factors[df_factors['dep'] == 'packing']['water_persec'].values[0]
        
        ####divider factors to get runtime of machine based on trx quan and product#####
        
        self.cotcard_div = 8
        self.polycard_div = 4
        self.draw_div = 6
        self.aco_div = 1824
        self.doubsix_div = 96
        self.doubeight_div =30
        self.doubten_div = 36
        self.twistsix_div = 1008
        self.twisteight_div = 312
        self.twistten_div = 300
        self.pencilwinder_div = 5
        
#%%
class costing(cno_det):
    
    
    def __init__(self, *args, **kwargs):
        super(costing, self).__init__(*args, **kwargs)
        
        
        ######all functions######
    def card_ecost(self):
        if self.abbrev in self.cotton_list:    
            return (self.cotcard_erate*self.trx_quan*self.card_explbs)        
        
        else:
            return (self.polycard_erate*self.trx_quan*self.card_explbs)
        
    def draw_ecost(self):
        return (self.draw_erate*self.trx_quan*self.draw_explbs)
    
    def aco_ecost(self):
        return (self.aco_erate*self.trx_quan*self.aco_explbs)

    def doub_ecost(self):
        if self.pkg_wt<7:
            return (self.doubsix_erate*self.trx_quan*self.doub_explbs)
        elif 7<=self.pkg_wt<9:
            return (self.doubeight_erate*self.trx_quan*self.doub_explbs)
        elif 9<= self.pkg_wt:
            return (self.doubten_erate*self.trx_quan*self.doub_explbs)

    def twist_ecost(self):
        if self.pkg_wt<7:
            return (self.twist_explbs*self.twistsix_erate*self.trx_quan)
        elif 7<=self.pkg_wt<9:
            return (self.twist_explbs*self.twisteight_erate*self.trx_quan)
        elif 9<= self.pkg_wt:
            return (self.twistten_erate*self.twist_explbs*self.trx_quan)
        
    def pencilwinder_ecost(self):
        return (self.pencilwind_explbs*self.pwt_erate*self.trx_quan)
    
    def cond_ecost(self):
        return (self.cond_erate*self.cond_explbs*self.trx_quan)
    
    def pack_ecost(self):
        return (self.pack_erate*self.shipping_explbs*self.trx_quan)
    
#    def backwinding_ecost(self):
#        return (self.backwind_explbs*self.backwinding_erate*self.trx_quan)

    def card_wcost(self):
        return (self.cotcard_watrate*self.trx_quan*self.card_explbs)
    
    
    def draw_wcost(self):
        return (self.draw_watrate*self.trx_quan*self.draw_explbs)
    
    def aco_wcost(self):
        return (self.aco_watrate*self.trx_quan*self.aco_explbs)

    def doub_wcost(self):
        if self.pkg_wt<7:
            return (self.doubsix_watrate*self.trx_quan*self.doub_explbs)
        elif 7<=self.pkg_wt<9:
            return (self.doubeight_watrate*self.trx_quan*self.doub_explbs)
        elif 9<= self.pkg_wt:
            return (self.doubten_watrate*self.trx_quan*self.doub_explbs)

    def twist_wcost(self):
        if self.pkg_wt<7:
            return (self.twist_explbs*self.twistsix_watrate*self.trx_quan)
        elif 7<=self.pkg_wt<9:
            return (self.twist_explbs*self.twisteight_watrate*self.trx_quan)
        elif 9<= self.pkg_wt:
            return (self.twistten_watrate*self.twist_explbs*self.trx_quan)
        
    def pencilwinder_wcost(self):
        return (self.pencilwind_explbs*self.pwt_watrate*self.trx_quan)
    
    def cond_wcost(self):
        return (self.cond_watrate*self.cond_explbs*self.trx_quan)
    
    def pack_wcost(self):
        return (self.pack_watrate*self.shipping_explbs*self.trx_quan)
    
    def total_ecost(self):
        return (self.card_ecost() + self.draw_ecost() + self.aco_ecost() + self.doub_ecost() + \
                self.twist_ecost() + self.pencilwinder_ecost() + self.cond_ecost() + self.pack_ecost())
        
    def total_wcost(self):
        return (self.card_wcost() + self.draw_wcost() + self.aco_wcost() + self.doub_wcost() + \
                self.twist_wcost() + self.pencilwinder_wcost() + self.cond_wcost() + self.pack_wcost())
        
        
    def total_cost(self):
        return (self.total_ecost() + self.total_wcost())
    
    def exp_time_month(self):
        card_time = self.card_explbs*self.trx_quan/self.cotcard_div # change to "if" condition
        draw_time = self.draw_explbs*self.trx_quan/self.draw_div
        aco_time = self.aco_explbs*self.trx_quan/self.aco_div
        doub_time = self.doub_explbs*self.trx_quan/self.doubsix_div # changeto "if condition
        twist_time = self.twist_explbs*self.trx_quan/self.twistsix_div #change to "if condition"
        pencilwinder_time = self.pencilwind_explbs*self.trx_quan/self.pencilwinder_div
        
        return card_time, draw_time, aco_time, doub_time, twist_time, pencilwinder_time
    
#%%
cost_list = []
cost_list_perlb = []

columns_time = ['cno', 'trx_quan', 'card_time', 'draw_time', 'aco_time', 'doub_time', 'twist_time', 'pencilwinder_time']

columns_dfcost = ['cno', 'trx_quan', 'card_ecost', 'card_wcost', 'draw_ecost', 'draw_wcost', 'aco_ecost', 'aco_wcost', \
                  'doub_ecost', 'doub_wcost', 'twist_ecost', 'twist_wcost', 'pencilwinder_ecost', 'pencilwinder_wcost', 'conditioning_ecost', \
                  'conditioning_wcost', 'pack_ecost', 'total_ecost', 'total_wcost', 'total_cost']

#%%

for i , row in df_cnodet.iterrows():
    cno_current = df_cnodet.iloc[i]['cno']
    print(cno_current)
    
    df_currentcno = df_cnodet[df_cnodet['cno'] == cno_current]
    costdata = costing(cno_current, df_currentcno, df_factors)
    print(costdata.cno)

    costx = pd.DataFrame([[costdata.cno, costdata.trx_quan, costdata.card_ecost(), costdata.card_wcost(),\
                           costdata.draw_ecost(), costdata.draw_wcost(), costdata.aco_ecost() , costdata.aco_wcost() ,\
                           costdata.doub_ecost(), costdata.doub_wcost(), costdata.twist_ecost(), \
                           costdata.twist_wcost(), costdata.pencilwinder_ecost(), costdata.pencilwinder_wcost(), \
                           costdata.cond_ecost(), costdata.cond_wcost(), costdata.pack_ecost(), costdata.total_ecost(),\
                           costdata.total_wcost(), costdata.total_cost()]], columns = columns_dfcost)
    
    print(costx)
    cost_list.append(costx)
    
    time_month = costdata.exp_time_month()
    
    

df_cost = pd.concat(cost_list)

#df_timemonth = pd.concat()
#%%

df_cost.to_csv (r'\\shufordyarnsllc.local\SYDFS\UserFiles$\SankalpMishra\Desktop\Sankalp_all\cost_assistanceproject_8_19_2019\Hickory Spinners\mar_2020\hs_mar20_cost.csv', index = None, header=True) 

    

