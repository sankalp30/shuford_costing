# -*- coding: utf-8 -*-
"""
Created on Tue Mar 10 09:13:40 2020

@author: SankalpMishra
"""

# -*- coding: utf-8 -*-
"""


@author: SankalpMishra
"""

import numpy as np
import pandas as pd

#%%
path = '//shufordyarnsllc.local/SYDFS/UserFiles$/SankalpMishra/Desktop/Sankalp_all/cost_assistanceproject_8_19_2019'

path_ds_all = '//shufordyarnsllc.local/SYDFS/UserFiles$/SankalpMishra/Desktop/Sankalp_all'
df_cno_trx = pd.read_excel(path + '/Aug_2020/' + 'Aug20_pp.xlsx', sheet_name = 'aug_pp',header = 0)
df_cno_trx['cno'] = df_cno_trx['Item Number'].map(lambda x: str(x)[:-6])
df_cno_trx['trx_quan'] = df_cno_trx['TRX QTY']
df_costpersec = pd.read_excel(path + '/input_tbls/' + 'current_utility_factor_DS-waterratetrial.xlsx', sheet_name = 'Updated_dep_factors', header = 0)

df_standards = pd.read_excel(path + '/Aug_2020/' + 'aug20_stds.xlsx', header = 0)

#df_cnodet = pd.read_excel(path + '/Feb2020/' + 'feb_20_allstd.xlsx', header = 0)

df_cnodet = df_standards.merge(df_cno_trx, on = ['cno'])
#%%
df_cnodet = df_cnodet.fillna(0)
#df_standards['cno'] = df_standards['cno'].astype('int64')


#%%
class cno_det:
    def __init__(self, current_cno, df_currentcno = [], df_costpersec = []):


        c= 1
        self.kwh_rate = (0.0456/3600)
        self.current_cno = current_cno
        self.df_currentcno = df_currentcno
        self.cno = df_currentcno['cno'].values[0]
        #self.ct = df_currentcno['ct/pyn/fpl'].values[0]
        self.preb_explbs = (df_currentcno['preblend'].values[0])
        self.card_explbs = df_currentcno['carding'].values[0]
        self.passes = df_currentcno['passes'].values[0]
        #self.spin_type = df_currentcno['spin_type'].values[0]
        self.draw_explbs = df_currentcno['drawing'].values[0]/c
        self.roving_explbs = df_currentcno['roving'].values[0]/c
        self.marzoli_explbs = df_currentcno['Marzoli_spin'].values[0]/c
        self.marzoli_wind = df_currentcno['Marzoli_wind'].values[0]/c
        self.zinser_explbs = df_currentcno['Zinser_spin'].values[0]/c
        self.zinser_wind = df_currentcno['Zinser_wind'].values[0]/c
        self.se_explbs = df_currentcno['SE9'].values[0]/c
        #self.oe_explbs = df_currentcno['oe_explbs'].values[0]/c
        self.aj_explbs = df_currentcno['Airjet'].values[0]/c #not lbs per hour but lbs per shift
        self.aco_explbs = df_currentcno['ACO'].values[0]/c
        self.doub_explbs = df_currentcno['Doubling'].values[0]/c
        self.pack_explbs = df_currentcno['Shipping'].values[0]/c
        self.packing_time = 4.2
        
        
        
        ####should be changed. remove 0.70 multipler and use standards lbs for subsequent runs
        self.twisting_explbs = df_currentcno['Twisting'].values[0] # incorrectly labelled. its lbs per shift and shouldn't be multplied by 8.100% lbs taken, using 70% efficency fixed
        ######
        
        
        self.duro_explbs = df_currentcno['Duro'].values[0]/c
        self.xeno_explbs = df_currentcno['Xeno'].values[0]/c
        self.cond_explbs = df_currentcno['Conditioning'].values[0]/c
        self.trx_quan = df_currentcno['trx_quan'].values[0]
        #self.pkg_length = df_currentcno['pkg_length'].values[0]
        self.pkg_wt = df_currentcno['pkg wt'].values[0]
        
        self.preb_elecrate = df_costpersec[df_costpersec['dep'] == 'PreBlend']['elect_persec'].values[0]
        self.card_elecrate = df_costpersec[df_costpersec['dep'] == 'Opening &Carding']['elect_persec'].values[0]
        self.draw_aj_elecrate = df_costpersec[df_costpersec['dep'] == 'Drawing']['elect_persec'].values[0]
        self.draw_oe_elecrate = df_costpersec[df_costpersec['dep'] == 'Drawing']['elect_persec'].values[0]
        self.draw_rs_elecrate = df_costpersec[df_costpersec['dep'] == 'Drawing']['elect_persec'].values[0]
        self.rov_elecrate = df_costpersec[df_costpersec['dep'] == 'Roving']['elect_persec'].values[0]
        self.aj_elecrate = df_costpersec[df_costpersec['dep'] == 'AJ']['elect_persec'].values[0]
        #self.oe_elecrate = df_costpersec[df_costpersec['dep'] == 'OE']['elect_persec'].values[0]
        
        self.aco_elecrate = df_costpersec[df_costpersec['dep'] == 'Aco8']['elect_persec'].values[0]
        self.se_elecrate = df_costpersec[df_costpersec['dep'] == 'Se9']['elect_persec'].values[0]
        
        self.rs_elecrate = df_costpersec[df_costpersec['dep'] == 'Ring']['elect_persec'].values[0] # extra space in spintype name
        
        self.marzolispin_elecrate = df_costpersec[df_costpersec['dep'] == 'Marzoli_spin']['elect_persec'].values[0]
        self.marzoliwind_elecrate = df_costpersec[df_costpersec['dep'] == 'Marzoli_wind']['elect_persec'].values[0]
        self.zinserspin_elecrate = df_costpersec[df_costpersec['dep'] == 'Zinser_spin']['elect_persec'].values[0]
        self.zinserwind_elecrate = df_costpersec[df_costpersec['dep'] == 'Zinser_wind']['elect_persec'].values[0]
        
        self.doubsix_elecrate = df_costpersec[df_costpersec['dep'] == 'Doubling -6 inch']['elect_persec'].values[0]
        self.doubeight_elecrate = df_costpersec[df_costpersec['dep'] == 'Doubling - 8 & 10 inch']['elect_persec'].values[0]
        self.twistsix_elecrate = df_costpersec[df_costpersec['dep'] == 'Twisting -6 inch']['elect_persec'].values[0]
        self.twisteight_elecrate = df_costpersec[df_costpersec['dep'] == 'Twisting - 8 & 10 inch']['elect_persec'].values[0]
        self.xeno_elecrate = df_costpersec[df_costpersec['dep'] == 'Xeno']['elect_persec'].values[0]
        self.duro_elecrate = df_costpersec[df_costpersec['dep'] == 'Duro']['elect_persec'].values[0]
        self.cond_elecrate = df_costpersec[df_costpersec['dep'] == 'Conditioning']['elect_persec'].values[0]
        self.pack_elecrate = df_costpersec[df_costpersec['dep'] == 'packing']['elect_persec'].values[0]
        
        
        self.preb_watrate = df_costpersec[df_costpersec['dep'] == 'PreBlend']['water_persec'].values[0]
        self.card_watrate = df_costpersec[df_costpersec['dep'] == 'Opening &Carding']['water_persec'].values[0]
        self.draw_aj_watrate = df_costpersec[df_costpersec['dep'] == 'Drawing']['water_persec'].values[0]
        self.draw_oe_watrate = df_costpersec[df_costpersec['dep'] == 'Drawing']['water_persec'].values[0]
        self.draw_rs_watrate = df_costpersec[df_costpersec['dep'] == 'Drawing']['water_persec'].values[0]
        self.rov_watrate = df_costpersec[df_costpersec['dep'] == 'Roving']['water_persec'].values[0]
        self.aj_watrate = df_costpersec[df_costpersec['dep'] == 'AJ']['water_persec'].values[0]
        #self.oe_watrate = df_costpersec[df_costpersec['dep'] == 'OE']['water_persec'].values[0]
        
        self.aco_watrate = df_costpersec[df_costpersec['dep'] == 'Aco8']['water_persec'].values[0]
        self.se_watrate = df_costpersec[df_costpersec['dep'] == 'Se9']['water_persec'].values[0]
        
        #self.rs_watrate = df_costpersec[df_costpersec['dep'] == 'Ring ']['water_persec'].values[0] # extra space in spintype name
        
        self.marzolispin_watrate = df_costpersec[df_costpersec['dep'] == 'Marzoli_spin']['water_persec'].values[0]
        self.marzoliwind_watrate = df_costpersec[df_costpersec['dep'] == 'Marzoli_wind']['water_persec'].values[0]
        self.zinserspin_watrate = df_costpersec[df_costpersec['dep'] == 'Zinser_spin']['water_persec'].values[0]
        self.zinserwind_watrate = df_costpersec[df_costpersec['dep'] == 'Zinser_wind']['water_persec'].values[0]
        
        self.doubsix_watrate = df_costpersec[df_costpersec['dep'] == 'Doubling -6 inch']['water_persec'].values[0]
        self.doubeight_watrate = df_costpersec[df_costpersec['dep'] == 'Doubling - 8 & 10 inch']['water_persec'].values[0]
        self.twistsix_watrate = df_costpersec[df_costpersec['dep'] == 'Twisting -6 inch']['water_persec'].values[0]
        self.twisteight_watrate = df_costpersec[df_costpersec['dep'] == 'Twisting - 8 & 10 inch']['water_persec'].values[0]
        self.xeno_watrate = df_costpersec[df_costpersec['dep'] == 'Xeno']['water_persec'].values[0]
        self.duro_watrate = df_costpersec[df_costpersec['dep'] == 'PreBlend']['water_persec'].values[0]
        self.cond_watrate = df_costpersec[df_costpersec['dep'] == 'Conditioning']['water_persec'].values[0]
        
        self.packing_watrate =  0 #df_costpersec[df_costpersec['dep'] == 'Conditioning']['water_persec'].values[0]
        
        


#%%
class costing(cno_det):
    def __init__(self, *args, **kwargs):
        super(costing, self).__init__(*args, **kwargs)
        
        
        
        
    def pkg_in_crate(self):
        if 0<self.pkg_wt<=3.5:
            return 216
        elif 3.5<self.pkg_wt<= 6:
            return 125
         
        elif 6 < self.pkg_wt <= 8:
            return 96
        
        elif 8<self.pkg_wt:
            return 64
    
        else:
            return 96
             
    
    def packing_explbs(self):
        return (self.pkg_in_crate()*self.pkg_wt*480/self.packing_time)/28800
     
        
    def preb_ecost(self):
        return ((self.preb_elecrate*self.trx_quan*self.preb_explbs) if self.preb_explbs>0 else 0)
    
    def preb_wcost(self):
        return ((self.preb_watrate*self.preb_explbs) if self.preb_explbs > 0 else 0)
    
    def card_ecost(self):
        return ((self.card_elecrate*self.trx_quan*self.card_explbs) if self.card_explbs > 0 else 0)
    
    def card_wcost(self):
        return ((self.card_watrate*self.trx_quan*self.card_explbs) if self.card_explbs > 0 else 0)
    
    def draw_ecost(self):
        return (self.draw_rs_elecrate*self.trx_quan*(self.draw_explbs) if (self.draw_explbs) >0 else 0)  
        
    def draw_wcost(self):
        return (self.draw_rs_watrate*self.trx_quan*(self.draw_explbs) if (self.draw_explbs) >0 else 0)
        
        
    def rov_ecost(self):
        return (self.rov_elecrate*self.trx_quan*self.roving_explbs if self.roving_explbs > 0 else 0)
    
    def rov_wcost(self):
        return (self.rov_watrate*self.trx_quan*self.roving_explbs if self.roving_explbs > 0 else 0)
        
#    def spin_ecost(self):
#        if self.spin_type == 'aj':
#            return ((self.aj_elecrate)*self.trx_quan*self.kwh_rate/self.aj_explbs if self.aj_explbs > 0 else 0)
#        elif self.spin_type == 'oe':
#            return (self.oe_elecrate*self.trx_quan*self.kwh_rate/(self.oe_explbs) if (self.oe_explbs)>0 else 0)
#        elif self.spin_type == 'rs':
#            return (self.rs_elecrate*self.trx_quan*self.kwh_rate/(self.marzoli_explbs + self.zinser_explbs) if (self.marzoli_explbs + self.zinser_explbs) > 0 else 0)
#        else:
#            return 0
        
    def marzoli_ecost(self):
        return ((self.marzolispin_elecrate)*self.trx_quan*self.marzoli_explbs if self.marzoli_explbs > 0 else 0) + \
                ((self.marzoliwind_elecrate)*self.trx_quan*self.marzoli_wind if self.marzoli_wind > 0 else 0)
                
    def marzoli_wcost(self):
        return ((self.marzolispin_watrate)*self.trx_quan*self.marzoli_explbs if self.marzoli_explbs > 0 else 0) + \
                ((self.marzoliwind_watrate)*self.trx_quan*self.marzoli_wind if self.marzoli_wind > 0 else 0)
    
    def zinser_ecost(self):
        return ((self.zinserspin_elecrate)*self.trx_quan*self.zinser_explbs if self.zinser_explbs > 0 else 0) + \
                ((self.zinserwind_elecrate)*self.trx_quan*self.zinser_wind if self.zinser_wind > 0 else 0)
                
    def zinser_wcost(self):
        return ((self.zinserspin_watrate)*self.trx_quan*self.zinser_explbs if self.zinser_explbs > 0 else 0) + \
                ((self.zinserwind_watrate)*self.trx_quan*self.zinser_wind if self.zinser_wind > 0 else 0)
    
    
    def aj_ecost(self):
        return ((self.aj_elecrate)*self.trx_quan*self.aj_explbs if self.aj_explbs > 0 else 0)
    
    def aj_wcost(self):
        return ((self.aj_watrate)*self.trx_quan*self.aj_explbs if self.aj_explbs > 0 else 0)
    
     
    def aco_ecost(self):
        return ((self.aco_elecrate)*self.trx_quan*self.aco_explbs if self.aco_explbs > 0 else 0)
    
    def aco_wcost(self):
        return ((self.aco_watrate)*self.trx_quan*self.aco_explbs if self.aco_explbs > 0 else 0)  
    
    
    def se_ecost(self):
        return ((self.se_elecrate)*self.trx_quan*self.se_explbs if self.se_explbs > 0 else 0)
    
    def se_wcost(self):
        return ((self.se_watrate)*self.trx_quan*self.se_explbs if self.se_explbs > 0 else 0)  
    
    
#    def spin_wcost(self):
#        if self.spin_type == 'aj':
#            return ((self.aj_watrate)*self.trx_quan/self.aj_explbs if self.aj_explbs > 0 else 0)
#        elif self.spin_type == 'oe':
#            return (self.oe_watrate*self.trx_quan/(self.oe_explbs) if (self.oe_explbs)>0 else 0)
#        elif self.spin_type == 'rs':
#            return (self.rs_watrate*self.trx_quan/(self.marzoli_explbs + self.zinser_explbs) if (self.marzoli_explbs + self.zinser_explbs) > 0 else 0)
#        else:
#            return 0
        
    def doub_ecost(self):
        if self.pkg_wt>=7:
            return (self.doubeight_elecrate*self.trx_quan*self.doub_explbs if self.doub_explbs>0 else 0)
        else:
            return (self.doubsix_elecrate*self.trx_quan*self.doub_explbs if self.doub_explbs>0 else 0)
        
    def doub_wcost(self):
        if self.pkg_wt>7:
            return (self.doubeight_watrate*self.trx_quan*self.doub_explbs if self.doub_explbs>0 else 0)
        else:
            return (self.doubsix_watrate*self.trx_quan*self.doub_explbs if self.doub_explbs>0 else 0)
        
    def twist_ecost(self):
        if self.pkg_wt > 7:
            return (self.twisteight_elecrate*self.trx_quan*self.twisting_explbs if self.twisting_explbs>0 else 0)
        else: 
             return (self.twistsix_elecrate*self.trx_quan*self.twisting_explbs if self.twisting_explbs > 0 else 0)
         
    def twist_wcost(self):
        if self.pkg_wt > 7:
            return (self.twisteight_watrate*self.trx_quan*self.twisting_explbs if self.twisting_explbs>0 else 0)
        else: 
             return (self.twistsix_watrate*self.trx_quan*self.twisting_explbs if self.twisting_explbs > 0 else 0)
                 
    def xeno_ecost(self):
        return (self.xeno_elecrate*self.trx_quan*self.xeno_explbs if self.xeno_explbs > 0  else 0)
    
    def xeno_wcost(self):
        return (self.xeno_watrate*self.trx_quan*self.xeno_explbs if self.xeno_explbs > 0  else 0)
    
    def duro_ecost(self):
        return (self.duro_elecrate*self.trx_quan*self.duro_explbs if self.duro_explbs > 0 else 0)
    
    def duro_wcost(self):
        return (self.duro_watrate*self.trx_quan*self.duro_explbs if self.duro_explbs > 0 else 0)
        
    def cond_ecost(self):
        return (self.cond_elecrate*self.trx_quan*self.cond_explbs if self.cond_explbs > 0 else 0)
    
    def cond_wcost(self):
        return (self.cond_watrate*self.trx_quan*self.cond_explbs if self.cond_explbs>0 else 0)
    
    def packing_ecost(self):
        return (self.pack_elecrate*self.trx_quan*self.pack_explbs if self.pack_explbs>0 else 0)
    
    def total_cost(self):
        return (self.preb_ecost()+ self.preb_wcost() + self.card_ecost() + self.card_wcost() + \
                self.draw_ecost() +  self.draw_wcost() + self.doub_ecost() + self.doub_wcost() + self.twist_ecost() + \
                self.twist_wcost() + self.xeno_ecost() + self.xeno_wcost() + self.duro_ecost() + self.duro_wcost() + self.cond_ecost() + self.cond_wcost() + self.packing_ecost() + \
                self.marzoli_ecost() + self.marzoli_wcost() + self.aj_ecost() + self.aj_wcost()+ self.aco_ecost() + self.aco_wcost() + self.se_ecost() + self.se_wcost() + \
                self.zinser_wcost() + self.zinser_ecost())
        
    
    def total_wcost(self):
        return (self.preb_wcost() + self.card_wcost() + \
                self.draw_wcost()  + self.rov_wcost() + self.doub_wcost() + \
                self.twist_wcost() +  self.xeno_wcost() +  self.duro_wcost() + self.cond_wcost() + \
                self.marzoli_wcost() + self.aj_wcost()+  self.aco_wcost() + self.se_wcost() + self.zinser_wcost())
    
    def total_ecost(self):
        return (self.preb_ecost() + self.card_ecost() + \
                self.draw_ecost() + self.rov_ecost()  +  self.doub_ecost() + \
                self.twist_ecost() +  self.xeno_ecost() +  self.duro_ecost() + self.cond_ecost() + self.packing_ecost()+ self.zinser_ecost() + \
                self.marzoli_ecost() + self.aj_ecost() + self.aco_ecost()  + self.se_ecost()) 
        
        
#%%
        
cost_list = []
#cost_list_perlb = []

columns_dfcost = ['const no', 'trx_qty', 'preb_ecost', 'preb_wcost', 'card_ecost', 'card_wcost', 'draw_ecost', 'draw_wcost', 'rov_ecost', 'rov_wcost', 'marzoli_ecost', 'zinser_ecost', 'aco_ecost', 'se_ecost', 'aj_ecost', \
           'marzoli_wcost', 'zinser_wcost', 'aco_wcost', 'se_wcost', 'doub_ecost', 'doub_wcost', 'twist_ecost', 'twist_wcost', 'xeno_ecost', 'xeno_wcost', 'duro_ecost', 'duro_wcost', 'cond_ecost', 'cond_wcost', 'pack_ecost','total_ecost', 'total_wcost', 'total_cost']

#columns_dfcost_lb = ['const no', 'trx_qty', 'spin_type', 'preb_ecost_perlb', 'preb_wcostperlb', 'card_ecostperlb', 'card_wcostperlb', 'draw_ecostperlb', 'draw_wcostperlb', 'rov_ecost_perlb', 'rov_wcost_perlb', 'spin_ecostperlb', \
#           'spin_wcostperlb', 'doub_ecostperlb', 'doub_wcostperlb', 'twist_ecostperlb', 'twist_wcostperlb', 'xeno_ecostperlb', 'xeno_wcostperlb', 'duro_ecostperlb', 'duro_wcostperlb', 'cond_ecostperlb', 'cond_wcostperlb','pack_ecostperlb','total_ecost_perlb', 'total_wcost_perlb', 'total_cost_perlb']

for i, row in df_cnodet.iterrows():
    cno_current = df_cnodet.iloc[i]['cno']
    print(cno_current)
    #cno_data = cno_det(current_cno, cno_det, df_costpersec)
    df_currentcno = df_cnodet[df_cnodet['cno'] == cno_current]
    costdata = costing(cno_current, df_currentcno, df_costpersec)
    print(costdata.cno, costdata.packing_explbs())
    costx = pd.DataFrame([[costdata.cno, costdata.trx_quan, costdata.preb_ecost(), costdata.preb_wcost(), costdata.card_ecost(), costdata.card_wcost(),\
                           costdata.draw_ecost(), costdata.draw_wcost(), costdata.rov_ecost(), costdata.rov_wcost(), costdata.marzoli_ecost(), costdata.zinser_ecost() , costdata.aco_ecost() , costdata.se_ecost(), costdata.aj_ecost() ,\
                           costdata.marzoli_wcost(), costdata.zinser_wcost() , costdata.aco_wcost() , costdata.se_wcost() ,\
                           costdata.doub_ecost(), costdata.doub_wcost(), costdata.twist_ecost(), \
                           costdata.twist_wcost(), costdata.xeno_ecost(), costdata.xeno_wcost(), costdata.duro_ecost(), costdata.duro_wcost(), \
                           costdata.cond_ecost(), costdata.cond_wcost(), costdata.packing_ecost(), costdata.total_ecost(),\
                           costdata.total_wcost(), costdata.total_cost()]], columns = columns_dfcost)
    
    print(costx)
    
    cost_list.append(costx)
    
#    costy = pd.DataFrame([[costdata.cno, costdata.trx_quan, costdata.spin_type, costdata.preb_ecost()/costdata.trx_quan, costdata.preb_wcost()/costdata.trx_quan, costdata.card_ecost()/costdata.trx_quan, costdata.card_wcost()/costdata.trx_quan,\
#                           costdata.draw_ecost()/costdata.trx_quan, costdata.draw_wcost()/costdata.trx_quan,  costdata.rov_ecost()/costdata.trx_quan, costdata.rov_wcost()/costdata.trx_quan, costdata.spin_ecost()/costdata.trx_quan, costdata.spin_wcost()/costdata.trx_quan, \
#                           costdata.doub_ecost()/costdata.trx_quan, costdata.doub_wcost()/costdata.trx_quan, costdata.twist_ecost()/costdata.trx_quan, 
#                           costdata.twist_wcost()/costdata.trx_quan, costdata.xeno_ecost()/costdata.trx_quan, costdata.xeno_wcost()/costdata.trx_quan, costdata.duro_ecost()/costdata.trx_quan, costdata.duro_wcost()/costdata.trx_quan, costdata.cond_ecost()/costdata.trx_quan, costdata.cond_wcost()/costdata.trx_quan, \
#                           costdata.packing_ecost()/costdata.trx_quan, costdata.total_ecost()/costdata.trx_quan,\
#                           costdata.total_wcost()/costdata.trx_quan, costdata.total_cost()/costdata.trx_quan]], columns = columns_dfcost_lb)
#    
#    cost_list_perlb.append(costy)
    
df_cost_feb2019 = pd.concat(cost_list)

#df_cost_feb2019_perlb = pd.concat(cost_list_perlb)
    
#%% Doesn't work, explicit close problem

#out_path = 'D:\cost_assistanceproject_8_19_2019\Feb2020\feb_2020_kwmodel_3_18_2020.xlsx'
#writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
#df_cost_feb2019.to_excel(writer, sheet_name = 'feb20_cnocost', index = False)
##df_cost_feb2019_perlb.to_excel(writer, sheet_name = 'feb19_cnocost_perlb', index = False)

#%% Use only when writing to csv!
df_cost_feb2019.to_csv (r'\\shufordyarnsllc.local\SYDFS\UserFiles$\SankalpMishra\Desktop\Sankalp_all\cost_assistanceproject_8_19_2019\Aug_2020\aug20_kwmodel_9_22_2020_overall.csv', index = None, header=True) 
