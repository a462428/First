import numpy as np
import math
import xlrd
import string
from random import *


def read_data(path,sheet_num):
    data = xlrd.open_workbook(path)
    table = data.sheets()[sheet_num]   #read excel sheet-0,1,2,3.....

    return table

            

def cal_similarity(table,item):
    similarity=[]
    
    #for j_a in range(1,table.ncols-1):
    for j_a in range(3,5):
        if(j_a==item):
            similarity.append(1)
            continue
        
        sum_ans_1=0
        sum_ans_2=0
        product=0

        sum_1 =0
        sum_2 =0
        num_1 =0
        num_2 =0
        
        for i in range(1,table.nrows):
            temp_1=table.cell(i,item).value
            temp_2=table.cell(i,j_a).value
            
            if( type(temp_1)==float  ):
                sum_1+=temp_1
                num_1+=1
            if( type(temp_2)==float  ):
                sum_2+=temp_2
                num_2+=1
        
        avg_1=1.0*sum_1/num_1
        avg_2=1.0*sum_2/num_2



        print table.nrows
        for i in range(1,table.nrows):
            temp_1=table.cell(i,item).value
            temp_2=table.cell(i,j_a).value
            if( type(temp_1)==float  ):
                sum_ans_1+=float( (temp_1-avg_1)*(temp_1-avg_1) )
            if( type(temp_2)==float  ):
                sum_ans_2+=float( (temp_2-avg_2)*(temp_2-avg_2) )
            if( type(temp_1)==float and type(temp_2)==float ):
                product+=float( (temp_1-avg_1)*(temp_2-avg_2) )
                
        similarity.append(  1.0*product / (math.sqrt(sum_ans_1*sum_ans_2) ))

    return np.array(similarity)





def range_similarity(similarity, table):
    
    rank = []
    rank.append( np.argsort(similarity) )  #+2
    print(similarity)
    print
    print(rank)
    print
    
    for i in range(len(rank[0])-2,len(rank[0])-7,-1):
        print(table.cell(0,rank[0][i]+1).value  )




def fill_score(table, people_id):
    score=[]
    for i in range(1,table.ncols-1):
        #temp=table.cell(people_id,i).value
        #if( type(temp)!=float ):
        similarity = cal_similarity(table, i)
        ans=find_score(table, people_id, i, similarity)
        score.append(ans)

    return np.array(score)
            
            
            

def find_score(table, people_id, num, similarity):
    ans=0
    sim_abs_sum=0
    sim_score_product=0
    sim_term=0
    
    for i in range(1,table.ncols-1):
        temp=table.cell(people_id,i).value
        #if( type(temp)==float):
        if(temp!=0):
            sim_score_product+=similarity[i-1]*temp if similarity[i-1]>0 else 0
            sim_abs_sum+=math.fabs( similarity[i-1]  )  if similarity[i-1]>0 else 0
            
    ans=1.0*sim_score_product / sim_abs_sum

    return ans




def range_movie(table,score):
    
    rank = []
    rank.append( np.argsort(score) )  #+2
    print
    for i in range(len(rank[0])-1,len(rank[0])-6,-1):
        print(rank[0][i],":",table.cell(0,rank[0][i]+1).value,":", score[rank[0][i]]  )
    
            
            

if __name__=='__main__':
    table = read_data('Assignment 3.xls',0)

    #ans=fill_score(table,2)
    #range_movie(table,ans)

    similarity=cal_similarity(table,3)
    
    print(similarity)
    #rank_name=range_similarity(similarity,table)
    
 


   
    
    

    
