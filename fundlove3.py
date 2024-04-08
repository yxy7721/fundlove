# -*- coding: utf-8 -*-
"""
Created on Wed Aug  2 16:50:50 2023

@author: yangxy
"""

import pandas as pd
import numpy as np
import copy
import seaborn as sns

class MakeTwoFactor:
    def __init__(self,dataset):
        self.adjfactor=dataset["adjfactor"]
        self.fundlove=dataset["fundlove"]
        self.stockclose=dataset["stockclose"]
        '''
        adjfactor=dataset["adjfactor"]
        fundlove=dataset["fundlove"]
        stockclose=dataset["stockclose"]
        '''
        
    def exe1(self):
        #为了形成因子
        for i,j in adjfactor.items():
            for k in j.values():
                fac=copy.deepcopy(k)
        for i,j in stockclose.items():
            for k in j.values():
                clo=copy.deepcopy(k)
        del i,j,k,adjfactor,stockclose
        datelist=set(fac["date"]) & set(clo["date"])
        datelist=pd.Series(list(datelist)).sort_values(ascending=True).reset_index(drop=True)
        clo=clo.set_index("date",drop=True)
        fac=fac.set_index("date",drop=True)
        clo=clo.stack(level=-1)
        fac=fac.stack(level=-1)
        clo=clo.reset_index(drop=0)
        fac=fac.reset_index(drop=0)
        clo["level_1"]=clo["level_1"].apply(lambda x:np.nan if ("HK" in x)or("BJ" in x)or("SZ" in x)or("688" in x) else x)
        clo=clo.dropna(axis=0,how="any")
        fac["level_1"]=fac["level_1"].apply(lambda x:np.nan if ("HK" in x)or("BJ" in x)or("SZ" in x)or("688" in x) else x)
        fac=fac.dropna(axis=0,how="any")
        clo=clo.set_index(["date","level_1"],drop=True)
        fac=fac.set_index(["date","level_1"],drop=True)
        df=pd.concat([clo,fac],axis=1)
        df.columns=["close","factor"]
        stocklist=df.index.get_level_values(level=1)
        stocklist=set(stocklist)-{"date"}
        stocklist=pd.Series(list(stocklist)).sort_values(ascending=True).reset_index(drop=True)
        stdate=pd.to_datetime("20180331")
        stdate=datelist[datelist>stdate].index[0]+14
        stdate=datelist[stdate]
        clo=clo.unstack(level=1)
        clo.columns=clo.columns.droplevel(level=0)
        fac=fac.unstack(level=1)
        fac.columns=fac.columns.droplevel(level=0)
        loveindex=pd.Series(fundlove.keys()).str.split(".",expand=True)[0]
        loveindex=pd.to_datetime(loveindex).sort_values(ascending=True)
        loveindex=pd.DataFrame(loveindex)
        loveindex["filename"]=pd.Series(fundlove.keys())
        df["f_mom"]="none"
        df["f_fundlove"]="none"
        df["l_ret"]="none"
        label_para={"rank":50,"future_days":30}
        for id_dt in range(len(datelist)):
            print(id_dt/len(datelist))
            dt=datelist[id_dt]
            if dt<stdate:
                continue
            if id_dt+label_para["future_days"]>len(datelist):
                break
            #df.loc[(dt,slice(None)),:][""]
            lag_1=df.loc[(datelist[id_dt+label_para["future_days"]],slice(None)),["close","factor"]]
            lag_0=df.loc[(datelist[id_dt],slice(None)),["close","factor"]]
            tmp1=lag_1.loc[:,"close"]*lag_1.loc[:,"factor"]
            tmp2=lag_0.loc[:,"close"]*lag_0.loc[:,"factor"]
            tmp1=copy.deepcopy(tmp1).reset_index(drop=False)
            tmp2=copy.deepcopy(tmp2).reset_index(drop=False)
            future_ret=copy.deepcopy(tmp1[0]/tmp2[0])
            future_ret.index=tmp1["level_1"]
            future_ret=future_ret.apply(lambda x:np.nan if x==np.inf else x)
            future_ret.dropna(inplace=True)
            future_ret=future_ret.sort_values(ascending=False).reset_index(drop=0)
            del tmp1,tmp2,lag_1,lag_0
            for code in stocklist:
                tmp1=datelist[datelist<dt].index[-1]
                tmp2=datelist[tmp1-59]
                tmp3=datelist[tmp1]
                tmp4=fac.loc[tmp3,code]*clo.loc[tmp3,code]
                tmp5=fac.loc[tmp2,code]*clo.loc[tmp2,code]
                if tmp4>0:
                    df.loc[(dt,code),"f_mom"]=tmp5/tmp4-1
                del tmp1,tmp2,tmp3,tmp4,tmp5
                tmp1=loveindex[dt<loveindex[0]].index[0]
                tmp1=loveindex.loc[tmp1,"filename"]
                tmp1=fundlove[tmp1]["file"]
                if len(tmp1[tmp1["代码"]==code].index)==0:
                    tmp2=50000
                    tmp3="nogood"
                else:
                    tmp2=tmp1[tmp1["代码"]==code].index[0]
                if tmp2<=50:
                    tmp3="50"
                elif tmp2<=100:
                    tmp3="100"
                elif tmp2<=200:
                    tmp3="200"
                elif tmp2<=300:
                    tmp3="300"
                elif tmp2<=500:
                    tmp3="500"
                elif tmp2<=10000:
                    tmp3="good"
                df.loc[(dt,code),"f_fundlove"]=tmp3
                del tmp1,tmp2,tmp3
                if len(future_ret[future_ret["level_1"]==code].index)==0:
                    continue
                tmp1=future_ret[future_ret["level_1"]==code].index[0]
                if label_para["rank"]>=tmp1:
                    df.loc[(dt,code),"l_ret"]=1
                else:
                    df.loc[(dt,code),"l_ret"]=0
                del tmp1

    def exe2(self):
        df.to_csv(r"D:\desktop\fundlove\df.csv",index=1,header=1)
        
    def exe3(self):
        df["l_ret"]=df["l_ret"].apply(lambda x:x if x!="none" else np.nan)
        df=df.dropna(axis=0,how="any")
        df["f_mom"]=df["f_mom"].apply(lambda x:x if x!="none" else np.nan)
        df=df.dropna(axis=0,how="any")
        df["f_fundlove"]=df["f_fundlove"].apply(lambda x:x if x!="none" else np.nan)
        df=df.dropna(axis=0,how="any")
        
        mldf=copy.deepcopy(df)
        from sklearn.preprocessing import OneHotEncoder
        onehot_encoder = OneHotEncoder(sparse_output=False)
        onehot_encoded = onehot_encoder.fit_transform(mldf[["f_fundlove"]])
        onehot_encoded=pd.DataFrame(onehot_encoded)
        onehot_encoded.columns=onehot_encoded.columns.map(lambda x:"f_fundlove"+"_"+str(x))
        onehot_encoded.index=mldf.index
        mldf=pd.concat([mldf,onehot_encoded],axis=1)
        mldf.drop(labels=["f_fundlove"],axis=1,inplace=True)
        mldf.drop(labels=["close","factor"],axis=1,inplace=True)
        
        
    def exe4(self):
        from sklearn.ensemble import GradientBoostingClassifier
        from sklearn.model_selection import train_test_split
        datelist=set(df.index.get_level_values(level=0))
        datelist=pd.Series(list(datelist)).sort_values(ascending=True)
        datelist=datelist.reset_index(drop=1)
        
        f=dict()
        for i in range(1,len(datelist)):
            dt=datelist[i]
            the_train=mldf.loc[(datelist[i-1],slice(None)),slice(None)]
            the_train=copy.deepcopy(the_train)
            the_train=the_train.droplevel(level=0)

            X_train, y_train = the_train.drop(['l_ret'],axis=1), the_train['l_ret']

            gbc=GradientBoostingClassifier(random_state=1)
            gbc.fit(X_train, y_train)
            
            tmp1=pd.Series(gbc.feature_importances_,index=the_train.drop(['l_ret'],axis=1).columns)
            tmp1=tmp1.sort_values(ascending=False)

            the_now=mldf.loc[(dt,slice(None)),slice(None)]
            the_now["predict"]=gbc.predict(the_now.drop(["l_ret"],axis=1))
            f[dt]=the_now[the_now["predict"]==1].index
            
        out=pd.DataFrame(columns=["value"])
        for i in range(1,len(datelist)):
            dt=datelist[i]
            f[dt]
            if i==1:
                value=1
            else:
                selected=list(f[datelist[i-1]].get_level_values(level=1))
                value=out.loc[datelist[i-1],"value"]
                tmp1=0
                for code in selected:
                    tmp4=fac.loc[datelist[i-1],code]*clo.loc[datelist[i-1],code]
                    tmp5=fac.loc[dt,code]*clo.loc[dt,code]
                    tmp1=tmp1+tmp5/tmp4*1/(len(selected))
                value=value*tmp1
            out.loc[dt,"value"]=value
                

        out.to_excel(r'D:\desktop\fundlove\双因子.xlsx',header=True,index=True)









