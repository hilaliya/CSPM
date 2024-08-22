# -*- coding: utf-8 -*-
"""
Created on Wed Jan 12 10:58:28 2022

@author: Hilal YILMAZ
"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from collections import namedtuple
from docplex.mp.model import Model

def cspm(loc,obj,time, **kwargs):
    
    # Open Workbook and worksheet
    wb = load_workbook(loc, keep_vba=True, data_only=True)
    ws = wb.active
    #print("Sheet name: ", ws)

    # Access Data
    n =  ws["B7"].value          # of charging stations
    Bmin=ws["B5"].value/100     # Min SOC level (%)
    l =  ws["B6"].value          # route length (kWh): same as the energy required to reach the destination (e_n+1)
    T =  ws["B8"].value          # Non-stop travel time: : same as the time required to reach the destination (s_n+1)
    B =  ws["B9"].value          # energy capacity of EV (kWh)
    r0 = ws["B10"].value        # remaining energy at origin (kWh)
    bp = ws["B14"].value/100   # Nonlinear charging time/brakepoints (Energy-time)
    sp = ws["B15"].value/100    # Nonlinear charging time/slopes (Energy-time)

    CS = []  # Charging stations
    for i in range(7, 7+n):
        cs = []
        for j in range(4, 9):
            col = get_column_letter(j)
            cs.append(ws[col+str(i)].value)
        cs = tuple(cs)
        CS.append(cs)
    # print(CS)   
        
    #Insert origin and destination nodes as dummy CSs
    CS.insert(0,(0,0,0,0,0))
    CS.insert(n+1,(n+1,l,T,0,0))
    
    # Sets
    CS_info = namedtuple("CS", ["order", "energy", "time","power","cost"])    
    cs = [CS_info(*i) for i in CS]
    cs = {i: cs[i] for i in range(0, n+2)}

    # Model
    mdl = Model(name="cspm", **kwargs)
    if time: mdl.set_time_limit(time)

    # Binary Decision Variables
    x = {i: mdl.binary_var(name="x%d" % (i)) for i in range(0, n+2)}
    z = {(i, j): mdl.binary_var(name="z%d.%d" % (i, j)) for i in range(1, n+2) for j in range(i)}
    b = {(i, j): mdl.binary_var(name="b%d.%d" % (i, j)) for i in range(1, n+2) for j in range(i)}
      
    # Continuous Decision Variables
    Ra = {i: mdl.continuous_var(ub=B, lb=B*Bmin, name="Ra%d" % (i)) for i in range(1, n+2)}
    Rd = {i: mdl.continuous_var(ub=B, lb=B*Bmin, name="Rd%d" % (i)) for i in range(0, n+1)}
    Ta = {i: mdl.continuous_var(lb=0, name="Ta%d" % (i)) for i in range(1, n+2)}
    Td = {i: mdl.continuous_var(lb=0, name="Td%d" % (i)) for i in range(0, n+1)}
     
    # Objective Function
    if obj== "traveltime": 
        objective = Ta[n+1]
    elif obj== "cost":
        objective = mdl.sum(cs[i].cost*(Td[i]-Ta[i]) for i in range(1, n+1))

    elif obj=="number_of_stops":
        objective = mdl.sum(x[i] for i in range(1, n+1))
                    
    if obj=="multi":
        mdl.set_multi_objective("min",[mdl.sum(cs[i].cost*(Rd[i]-Ra[i]) for i in range(1, n+1)), Ta[n+1]])    
    else:
        mdl.minimize(objective)

    
    # K1 (4): EV must be charged at a CS that can be reached with the initial remaining energy (r0)
    lhs1 = [cs[i].order for i in cs if cs[i].energy <= r0-B*Bmin and i>0]
    k1   = mdl.add_constraint(mdl.sum(x[i] for i in lhs1) >= 1)
    #print(k1)

    # K2 (5): EV must be charged at a CS that requires energy less than the energy capacity (B) to reach the destination
    lhs2 = [cs[i].order for i in cs if cs[n+1].energy-cs[i].energy <= B*(1-Bmin) and i<=n]
    k2   = mdl.add_constraint(mdl.sum(x[i] for i in lhs2) >= 1)
    #print(k2)

    # K3 (6-7): If the EV is charged both at i and j, z equals to 1
    for i in range(2, n+2):
        for j in range(1,i):
            k3  = mdl.add(x[i]+x[j]-1<=z[i,j])
            k3_ = mdl.add(x[i]+x[j]>=2*z[i,j])
            #print(i, j, k3)

    # K4 (8-9): If the EV is charged at CS between z[i,j]=1 and z[j,k] =1, then b=1
    for i in range(2, n+2):
        for k in range(1,i):
            k4  = mdl.add(mdl.sum(z[i,j] for j in range(k+1,i))<=n*b[i,k])
            k4_ = mdl.add(mdl.sum(z[i,j] for j in range(k+1,i))>=b[i,k])

    #K5 (10): The energy required to travel between two consecutive CSs  must not exceed the energy of the EV      
    for i in range(1, n+2):
        for j in range(1,i):           
            k5 = mdl.add(z[(i, j)]-b[(i, j)]*(cs[i].energy-cs[j].energy) <= Rd[j]-Ra[i])
 
    #K6 (11): Energy Balance                  
    for i in range(1, n+2):
        k6=mdl.add(cs[i].energy-cs[i-1].energy==Rd[i-1]-Ra[i])             


    #K7 (12): Time Balance                  
    for i in range(1, n+2):
        k7=mdl.add(cs[i].time-cs[i-1].time==Ta[i]-Td[i-1])
 
    
    #K8 (13): Piecewise Linear Function: f(desired energy)=time (min) needed for charging to reach desired energy
    for i in range(1,n+1):
        ctime=mdl.piecewise(0, [(0,0),(B*bp,60*B*bp/cs[i].power)], 60/(cs[i].power*sp)) #the piecewise function     
        k8=mdl.add(Td[i]-Ta[i]==ctime(Rd[i])-ctime(Ra[i])) 
        
    #K9 (14)  # included as a bound when defining the variables
    # for i in range(1,n+1):
    #     k9 = mdl.add(Ra[i]>=B*Bmin*x[i])  
     
    #K10 (15-18)
    for i in range(1,n+1):
        kq = mdl.add(mdl.if_then(x[i]==0,Rd[i]-Ra[i]==0))
        kq = mdl.add(mdl.if_then(x[i]==0,Td[i]-Ta[i]==0))


    #K14 (19): 
    mdl.add(x[0]==1)
    
    #K15 (20): 
    mdl.add(x[n+1]==1)
    
    #K16 (21): 
    mdl.add(Rd[0]==r0)
    
    #K17 (22): 
    mdl.add(Td[0]==0)

    
    tottime=Ta[n+1]
    mdl.add_kpi(tottime, publish_name="Travel time")
    
    totcost=mdl.sum(cs[i].cost*(Td[i]-Ta[i]) for i in range(1, n+1))
    mdl.add_kpi(totcost, publish_name="Cost")
       
    return mdl

