import numpy as np
import math
import xlrd
import string
import pandas as pd
from pandas import DataFrame,Series
from random import *


def read_data(path):
    table_0 = pd.read_excel("Assignment 6.xlsx",sheetname=0)
    table_1 = pd.read_excel("Assignment 6.xlsx",sheetname=1)

    return table_0, table_1
    


def get_svd(table):
    
    s= table.iloc[0:1,:]
    S = np.diag( s.values[0])

    v = table_0.iloc[2:,:]
    V = v.values

    u = table_1.iloc[0:,1:]
    U = u.values
    
    return U, S, V



            

if __name__=='__main__':

    table_0, table_1 = read_data("Assignment 6.xlsx")
 
    U , S , V= get_svd(table_0)
    
    score=np.dot( U, np.dot(S, np.transpose(V)) )

    #print table_0.iloc[np.argsort(score[8])[98]+2,0:0]
    for i in range(0,5):
        print table_0.iloc[np.argsort(score[8])[99-i]+2 ,0:0]

 

   
    
    

    
