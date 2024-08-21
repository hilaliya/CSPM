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
    
    # Connect to Excel Data
    #loc = ("rclm/test%d.xlsx" %(test)) #("toy-problem.xlsx")

    # Open Workbook and worksheet
    wb = load_workbook(loc, keep_vba=True, data_only=True)
    ws = wb.active
    #print("Sheet name: ", ws)

    # Access Data
    alt=ws["B5"].value/100      # Min SOC bound
    l = ws["B6"].value          # route length
    n = ws["B7"].value          # of charging stations
    T = ws["B8"].value          # Non-stop travel time
    B = ws["B9"].value          # energy capacity of EV
    r0 = ws["B10"].value        # remaining energy at origin  
    bp= [ws["B14"].value/100]   # Nonlinear charging time/brakepoints (Energy-time)
    sp=[ws["B15"].value/100]    # Nonlinear charging time/slopes (Energy-time)
    print("B:", B)
    CS = []  # Charging stations
    k = 0
    for i in range(7, 7+n):
        cs = []
        for j in range(4, 9):
            col = get_column_letter(j)
            cs.append(ws[col+str(i)].value)
        cs = tuple(cs)
        CS.append(cs)
    # print(CS)   
        
    #Insert origing and destination as dummy CSs
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
    Ra = {i: mdl.continuous_var(ub=B, lb=0, name="Ra%d" % (i)) for i in range(1, n+2)}
    Rd = {i: mdl.continuous_var(ub=B, lb=0, name="Rd%d" % (i)) for i in range(0, n+1)}
    Ta = {i: mdl.continuous_var(lb=0, name="Ta%d" % (i)) for i in range(1, n+2)}
    Td = {i: mdl.continuous_var(lb=0, name="Td%d" % (i)) for i in range(0, n+1)}
  
    
    # Objective Function
    if obj== "traveltime": 
        objective = Ta[n+1]
    elif obj== "cost":
        objective = mdl.sum(cs[i].cost*(Rd[i]-Ra[i]) for i in range(1, n+1))

    elif obj=="mincs":
        objective = mdl.sum(x[i] for i in range(1, n+1))
            
        
    if obj=="multi":
        mdl.set_multi_objective("min",[mdl.sum(cs[i].cost*(Rd[i]-Ra[i]) for i in range(1, n+1)), Ta[n+1]])    
    else:
        mdl.minimize(objective)


    #K0: 
    mdl.add(x[0]==1)
    mdl.add(x[n+1]==1)
    mdl.add(Rd[0]==r0)
    mdl.add(Td[0]==0)
    # K1: EV must be charged at a CS that can be reached with the initial remaining energy (r0)
    lhs1 = [cs[i].order for i in cs if cs[i].energy <= r0-B*alt and i>0]
    k1 = mdl.add_constraint(mdl.sum(x[i] for i in lhs1) >= 1)
    #print(k1)

    # K2: EV must be charged at a CS that requires energy less than the energy capacity (B) to reach the destination
    lhs2 = [cs[i].order for i in cs if l-cs[i].energy <= B*(1-alt) and i<=n]
    k2 = mdl.add_constraint(mdl.sum(x[i] for i in lhs2) >= 1)
    #print(k2)

    # K3: If the EV is charged both at i and j, z equals to 1
    for i in range(1, n+2):
        for j in range(i):
            k3 = mdl.add(x[i]+x[j]-1<=z[i,j])
            k3_= mdl.add(x[i]+x[j]>=2*z[i,j])
            #print(i, j, k3)

    # K4: If the EV is charged at CS between z[i,j]=1 and z[j,k] =1, then b=1
    for i in range(1, n+2):
        for k in range(i):
            k4 = mdl.add(mdl.sum(z[i,j] for j in range(k+1,i))>=b[i,k])
            k4_= mdl.add(mdl.sum(z[i,j] for j in range(k+1,i))<=n*b[i,k])

    #K5: The energy required to travel between two consecutive CSs  must not exceed the energy of the EV 
            
    for i in range(1, n+2):
        for j in range(i):           
            k7 = mdl.add(mdl.if_then(z[(i, j)]-b[(i, j)]==1,cs[i].energy-cs[j].energy == Rd[j]-Ra[i]))
            #k7 = mdl.add(z[(i, j)]-b[(i, j)]*(cs[i].energy-cs[j].energy) <= Rd[j]-Ra[i])
 
             
    # K8-10: Time Balance    
                 
    for i in range(1, n+2):
        k10=mdl.add(cs[i].time-cs[i-1].time==Ta[i]-Td[i-1])
 
    
    #K11: Piecewise Linear Function: f(desired energy)=time (min) needed for charging to reach desired energy
    for i in range(1,n+1):
        ctime=mdl.piecewise(0, [(0,0),(B*bp[0],60*B*bp[0]/cs[i].power)], 60/(cs[i].power*sp[0]))

        #Total charging time
        k11=mdl.add(Td[i]-Ta[i]==ctime(Rd[i])-ctime(Ra[i]))
        

    for i in range(1,n+1):
        kq = mdl.add(Ra[i]>=B*alt*x[i])
        kq = mdl.add(mdl.if_then(x[i]==0,Rd[i]-Ra[i]==0))
        kq = mdl.add(mdl.if_then(x[i]==0,Td[i]-Ta[i]==0))
        #kq = mdl.add(Rd[i]-Ra[i]>=(cs[i+1].energy-cs[i].energy)*x[i])


    kq = mdl.add(Ra[n+1]>=B*alt)

    
    tottime=Ta[n+1]
    mdl.add_kpi(tottime, publish_name="Travel time")
    
    totcost=mdl.sum(cs[i].cost*(Rd[i]-Ra[i]) for i in range(1, n+1))
    mdl.add_kpi(totcost, publish_name="Cost")
       
    return mdl

