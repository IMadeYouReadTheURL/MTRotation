#!/usr/bin/env python3

import numpy as np
import scipy.optimize as scop
import matplotlib.pyplot as plt
from openpyxl import load_workbook as lw

def func(x, c, d, k, n):
    return  c + d*x**n / (k**n + x**n );


book=lw('merged results.xlsx')
ws=book["Sheet1"]
table=np.array([[cell.value for cell in col] for col in ws['B1':'J62']])
r,c = table.shape
avrate=np.zeros(4)
num=0
avrhoderate=np.zeros(4)
numrhode=0

for i in range(2,c):
    name=table[0,i]
    unydata=np.array(table[1:r,i].astype(np.float))
    unxdata=np.array(table[1:r,0].astype(np.float))
    n,m=np.argmin(unydata),np.argmax(unydata)
    if m<n:
        print(i,m,n)
        break
    ydata=np.array(table[n+1:m+1,i].astype(np.float))
    xdata=np.array(table[n+1:m+1,0].astype(np.float))
    popt, pcov= scop.curve_fit(func, xdata,ydata, p0=[6e+02,1.5e+3,1e+3,4],bounds=(0,np.inf))
    if "Rhod" in name:
        avrhoderate+=popt
        numrhode+=1
    else:
        avrate+=popt
        num+=1
    print(num, avrate,numrhode,avrhoderate)
    plt.figure()
    plt.title("Graph displaying flourescence against time for test %s"%(name))
    xxdata=np.linspace(xdata[0],xdata[-1],2000)
    plt.plot(xxdata,func(xxdata,popt[0],popt[1],popt[2],popt[3]), label="fitted function",color='red')
    if "Rhod" in name:
        leg="Rhodamine Labeled MTs"
    else:
        leg="Unlabelled Microtubles"
    plt.plot(unxdata,unydata, 'bs' , label="original data (" + str(leg) + ")")
    plt.xlabel('time (s)')
    plt.ylabel('Flouresence intensity')
    plt.legend(loc=0)
    plt.savefig("test %s.png"%(name))
avrate=avrate/(num)
avrhoderate=avrhoderate/(numrhode)
print(avrate, avrhoderate)