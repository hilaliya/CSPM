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
objective="number_of_stops"  #traveltime, cost, multi, number_of_stops
time=100
usetime=False
loc= "case.xlsx" 
wb = load_workbook(loc, keep_vba=False, data_only=True)
ws = wb.active
if usetime:
    mdl =modl(loc, objective, time)
else:
    mdl =modl(loc, objective)
wb.close()
    

mdl.print_information()

if mdl.solve():
    
    print("\n-----------------------------",
          "\n  ***We Have A Solution!***", "\n-----------------------------\n")
   # mdl.print_solution()

    print("Case study")
    print(mdl.solve_details)
    print("Solve time:", round(mdl.solve_details.time,2))
    print("Objective: ", round( mdl.objective_value,2))  
    print(mdl.report_kpis())
    
    with open("case_%s.txt" %(objective), "w") as solfile:
        solfile.write(mdl.solution.to_string())   
            
                
else:
    print("Problem could not be solved ")
    print(mdl.get_solve_details())

gc.collect()

