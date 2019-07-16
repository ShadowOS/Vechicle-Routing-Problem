# -*- coding: utf-8 -*-
"""
Created on Thu Apr  4 22:25:30 2019

@author: Karnika
"""

from gurobipy import*
import os
import xlrd

book = xlrd.open_workbook(os.path.join("vrp.xlsx"))



Node=[]
Demand={}

sh = book.sheet_by_name("demand")

i = 1
while True:
    try:
        sp = sh.cell_value(i,0)
        Node.append(sp)
        Demand[sp]=sh.cell_value(i,1)        
        i = i + 1
        
    except IndexError:
        break

cost={}
sh = book.sheet_by_name("Cost")
i = 1
for P in Node:
    j = 1
    for Q in Node:
        cost[P,Q] = sh.cell_value(i,j)
        j += 1
    i += 1
Aij = {}
sh = book.sheet_by_name("Aij")

i = 1
for P in Node:
    j = 1
    for Q in Node:
        Aij[P,Q] = sh.cell_value(i,j)
        j += 1
    i += 1
numberOfVechile=2
K=numberOfVechile

m=Model("VRP")

m.modelSense=GRB.MINIMIZE
xij=m.addVars(Node,Node,vtype=GRB.BINARY,name='X_ij')
m.setObjective(sum(cost[i,j]*xij[i,j] for i in Node for j in Node))

for i in Node:
    m.addConstr(sum(Aij[i,j]*xij[i,j] for j in Node)==1)
    
for i in Node:
    m.addConstr(sum(xij[j,i]*Aij[j,i] for j in Node)==1)

#for i in Node:
#    m.addConstr(sum(xij[0,j] for j in Node)==K)

#for i in Node:
#    for j in Node:
#        if i!=j:
#            m.addConstr(xij[i,j]+ xij[j,i]==1)

m.optimize()
                
m.write('VRP.lp')
                            
for v in m.getVars():
    if v.x > 0.01:
        print(v.varName, v.x)
print('Objective:',m.objVal)
            

    
