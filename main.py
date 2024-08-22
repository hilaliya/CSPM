# -*- coding: utf-8 -*-
"""
Created on Fri Apr  1 21:14:20 2022

@author: Hilal YILMAZ
"""



import gc
from CSPM import cspm
from openpyxl import load_workbook

model="cspm"    
modl=eval(model)
objective="cost"  #traveltime, cost, multi, number_of_stops
time=1200 #max solve time (seconds)
usetime=True
loc='test3.xlsx' #"case_study.xlsx" 
wb = load_workbook(loc, keep_vba=False, data_only=True)
ws = wb.active
mdl =modl(loc, objective, time)

wb.close()
    

mdl.print_information()

if mdl.solve():
    
    print("\n-----------------------------",
          "\n  ***We Have A Solution!***", "\n-----------------------------\n")
    #mdl.print_solution()
    print("Case study")
    print(mdl.solve_details)
    print("Solve time:", round(mdl.solve_details.time,2))
    print("Objective: ", round( mdl.objective_value,2))  
    print(mdl.report_kpis())
    
    with open(loc[:-5]+"_%s.txt" %(objective), "w") as solfile:
        solfile.write(mdl.solution.to_string())   
            
                
else:
    print("Problem could not be solved ")
    print(mdl.get_solve_details())

gc.collect()

