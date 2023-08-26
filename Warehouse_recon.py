import numpy as np
import pandas as pd

def warehouse_recon(filepath): 
    df = pd.read_excel(filepath, sheet_name=None) 
    df1 = df["logistics"] 
    df2 = df["warehouse"] 
    comparison_values = df1.eq(df2) 
    rows,cols=np.where(comparison_values==False)
    post = []
    Rec_report = open("C:/Users/840/Desktop/reconciliation_report.txt", "w")
    for i in range(len(rows)):
        post.append([rows[i], cols[i]])
    for i in range(len(post)):
        print(Rec_report.write("There is a difference of "
                               +repr(abs(df1.iloc[post[i][0], post[i][1]] - df2.iloc[post[i][0], post[i][1]]))+
                               " between logistics and warehouse stock for "+ repr(df1.iloc[post[i][0],0]) + repr(df1.columns[post[i][1]])+"\n"))
        df1.iloc[post[i][0], post[i][1]] = '{} --> {}'.format(df1.iloc[post[i][0], post[i][1]],df2.iloc[post[i][0], post[i][1]])
    Rec_report.close()
    print(df1.head())

warehouse_recon("C:/Users/840/Desktop/Real_compare.xlsx")