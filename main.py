# -*- coding: utf-8 -*-
"""
Created on Fri Apr  1 21:14:20 2022

@author: Hilal YILMAZ
"""

from CSPM import cspm

loc='toyproblem.xlsx'   # the directory of the problem data
objective="multi"       #traveltime, cost, multi, number_of_stops
time=1200               #max solve time (seconds)
usetime=True            #if True, the model gives a solution under a time limit specified in "time"


mdl =cspm(loc, objective, time)    
mdl.print_information()

if mdl.solve():
    
    print("\n-----------------------------",
          "\n  ***We Have A Solution!***", 
          "\n-----------------------------\n")
   #mdl.print_solution()
    print(mdl.solve_details)
    print("Solve time:", round(mdl.solve_details.time,2))
    print("Objective: ", round( mdl.objective_value,2), "\n")  
    mdl.report_kpis()
    
    with open(loc[:-5]+"_%s.txt" %(objective), "w") as solfile:
        solfile.write(mdl.solution.to_string())   
            
                
else:
    print("Problem could not be solved ")
    print(mdl.get_solve_details())
