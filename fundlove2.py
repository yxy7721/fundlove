# -*- coding: utf-8 -*-
"""
Created on Tue Apr  4 11:17:26 2023

@author: yangxy
"""


import pandas as pd
import numpy as np
import copy
import openpyxl as op
import os
import xlwings as xw
import datetime

class BasicData:
    def __init__(self,dirpath,app):
        self.app=app
        self.dirpath=dirpath

    def read_one_after_one(self,date=True):
        dirlist=os.listdir(self.dirpath)
        greatlis=dict()
        filenamelis=list()
        for filename in dirlist:
            print(filename)
            filenamelis.append(filename)
            if os.path.isdir(os.path.join(self.dirpath,filename)) or filename=="ok.xlsx" or ("~$" in filename):
                continue
            wb=self.app.books.open(os.path.join(self.dirpath,filename))
            #wb[wb.sheetnames[0]].title
            greatdf=dict()
            for she in range(len(wb.sheets)):
                wb.sheets[she].name
                wb.sheets[0].used_range.last_cell.row
                wb.sheets[0].used_range.last_cell.column
                df=wb.sheets[she].range('A1').options(pd.DataFrame,index=False,expand="table").value
                #视情况看是否要对齐至季度最后一天
                if date:
                    df.index=pd.to_datetime(df['date'])
                    del df['date']
                    #df=df.resample('3M',axis=0,closed="right",label="right").last()
                    df=df.reset_index(drop=False)
                greatdf[wb.sheets[she].name]=df
            greatlis[filename]=greatdf
            wb.close()
        return greatlis,filenamelis
    
    def just_pile(self,ept):
        sheetset=set()
        pile_up_dict=dict()
        ept_dict=dict()
        greatlis=copy.deepcopy(self.greatlis)
        for filename in greatlis.keys():
            sheetset=sheetset | set(greatlis[filename].keys())
        for filename in greatlis.keys():
            print(filename)
            if filename in ept:
                ept_dict[filename]=greatlis[filename].copy()
                continue
            #filename='stockclose2018.xlsx' 
            for sheetname in sheetset:
                if sheetname in pile_up_dict.keys():
                    tmp1=pile_up_dict[sheetname].copy()
                    tmp2=greatlis[filename][sheetname].copy()
                    pile_up_dict[sheetname]=pd.concat([tmp1,tmp2],axis=0)
                else:
                    pile_up_dict[sheetname]=greatlis[filename][sheetname].copy()
        for i in pile_up_dict.keys():
            pile_up_dict[i].reset_index(drop=True)
        pile_up_dict={"smooth":pile_up_dict}
        return pile_up_dict,ept_dict

    def get_data(self):
        self.greatlis,self.filenamelis=self.read_one_after_one()
        self.pile_up_dict,self.ept_dict=self.just_pile(["null"])
        return self.pile_up_dict,self.ept_dict

class FundloveData(BasicData):
    def get_data(self,date=True):
        self.greatlis,self.filenamelis=self.read_one_after_one(date=False)
        return self.greatlis,self.filenamelis
    
class PrepareData():
    def __init__(self,filenamelist,app):
        self.stockpath=filenamelist[0]
        self.adjpath=filenamelist[1]
        self.fundlovepath=filenamelist[2]
        self.app=app
    def calcu(self):
        close=BasicData(self.stockpath,self.app)
        stockclose_dict,_=close.get_data()
        
        adjfac=BasicData(r"D:\desktop\mydatabase\adjfac",self.app)
        adjfactor_dict,_=adjfac.get_data()
        
        fdlove=FundloveData(r"D:\desktop\mydatabase\fundlove",self.app)
        fundlove,_=fdlove.get_data()
        del close,adjfac,fdlove
        back={"stockclose":stockclose_dict,
              "adjfactor":adjfactor_dict,
              "fundlove":fundlove,
            }
        return back
    def make_data(self):
        return self.calcu()

class MyPortfolio1(dict):
    def __init__(self,datadict):
        self.stockclose_dict=datadict["stockclose"]
        self.adjfactor_dict=datadict["adjfactor"]
        self.fundlove=datadict["fundlove"]

    def init_price_and_factor(self):
        for i1 in self.stockclose_dict.values():
            for stockclose_sht in i1.values():
                break
        self.stockclose_sht=stockclose_sht.copy()
        self.stockclose_sht.index=stockclose_sht["date"]
        del self.stockclose_sht["date"]
        for i1 in self.adjfactor_dict.values():
            for adjfac_sht in i1.values():
                break
        self.adjfac_sht=adjfac_sht.copy()
        self.adjfac_sht.index=adjfac_sht["date"]

    def get_datelist(self,dictlist,level=1):
        def unravel(dictlist):
            newlist=list()
            for d in dictlist:
                for i in d.keys():
                    newlist.append(d[i])
            return newlist
        for i in range(level):
            dictlist=unravel(dictlist)
        dflist=dictlist.copy()
        datelist=set()
        for dfdict in dflist:
            pass
            if len(datelist)==0:
                for df in dfdict.values():
                    datelist=set(df["date"])
            else:
                for df in dfdict.values():
                    datelist=set(df["date"])
                datelist=set(df["date"]) & datelist
        datelist=pd.Series(list(datelist))
        datelist=datelist.sort_values(ascending=True).reset_index(drop=True)
        datelist=pd.to_datetime(datelist)
        datelist.name="date"
        return datelist
    
    def cal_datelist(self):
        self.datelist=self.get_datelist([self.adjfactor_dict,self.stockclose_dict])
        
class MyPortfolio2(MyPortfolio1):
    def __init__(self,datadict,num_of_stock):
        super().__init__(datadict)
        self.num_of_stock=num_of_stock
        
    def search_adjust_day(self,adjust_day):
        i=0
        while True:
            if adjust_day+datetime.timedelta(days=i) in self.datelist.values:
                break
            else:
                i=i+1
        adjust_day=adjust_day+datetime.timedelta(days=i)
        adjindex=self.datelist[self.datelist==adjust_day].index[0]
        adjust_day=self.datelist.loc[adjindex]
        return adjust_day
    
    def create_adjustdatelist(self):
        self.cal_datelist()
        self.adjdatelist=list()
        for i,fundlove_filename in enumerate(self.fundlove.keys()):
            today=pd.to_datetime(fundlove_filename.split(".")[0])
            if i==(len(self.fundlove.keys())-1):
                break
            adjust_day=self.search_adjust_day(today)
            self.adjdatelist.append([fundlove_filename,adjust_day])
        self.adjdatelist=pd.DataFrame(self.adjdatelist)
        self.adjdatelist.columns=["filename","adjday"]
    
    def init_price_and_factor(self):
        stockclose_dict=copy.deepcopy(self.stockclose_dict)
        adjfactor_dict=copy.deepcopy(self.adjfactor_dict)
        for i1 in stockclose_dict.values():
            for stockclose_sht in i1.values():
                break
        stockclose_sht=stockclose_sht.copy()
        stockclose_sht.index=stockclose_sht["date"]
        del stockclose_sht["date"]
        for i1 in adjfactor_dict.values():
            for adjfac_sht in i1.values():
                break
        adjfac_sht=adjfac_sht.copy()
        adjfac_sht.index=adjfac_sht["date"]
        self.adjfac_sht=copy.deepcopy(adjfac_sht)
        self.stockclose_sht=copy.deepcopy(stockclose_sht)
        
    def is_valid(self,trade_day,code1):
        if self.stockclose_sht.loc[trade_day,code1]*self.adjfac_sht.loc[trade_day,code1]==0:
            return False
        elif (self.stockclose_sht.loc[trade_day,code1]!=self.stockclose_sht.loc[trade_day,code1]):
            return False
        elif self.adjfac_sht.loc[trade_day,code1]!=self.adjfac_sht.loc[trade_day,code1]:
            return False
        else:
            return True   
    
    def calcu_MA(self,trade_day,code,period=30):
        if self.datelist[self.datelist==trade_day].index[0]<=period:
            idx1=0
        else:
            idx1=self.datelist[self.datelist==trade_day].index[0]-period
        dtlis=self.datelist.iloc[idx1:idx1+period]
        MA1=(self.stockclose_sht.loc[dtlis,code]*self.adjfac_sht.loc[dtlis,code])
        MA1=MA1.map(lambda x:np.nan if x==0.0 else x).mean(skipna=True)
        return MA1

    def calcu_daily_holdings(self):
        self.init_price_and_factor()
        self.create_adjustdatelist()
        del self.stockclose_dict,self.adjfactor_dict
        
        my_portfolio1=dict()
        for i in range(len(self.datelist)):
        #for i in range(115):
            #i=74,115
            trade_day=self.datelist[i]
            if trade_day<self.adjdatelist["adjday"][0]:
                continue
            print(i/len(self.datelist))            
            j=0
            while True:
                if (trade_day<self.adjdatelist["adjday"][j+1]) and (self.adjdatelist["adjday"][j]<=trade_day):
                    break
                else:
                    j=j+1
                if j>=len(self.adjdatelist.index)-1:
                    break
            filename=self.adjdatelist["filename"][j]
            love_sheet=list(self.fundlove[filename].values())[0]
            selected_index=list()
            
            for j in range(len(love_sheet.index)):
                code=love_sheet.loc[j,"代码"]
                if (code in self.stockclose_sht.columns) and (code in self.adjfac_sht.columns):
                    pass
                else:
                    continue
                if self.is_valid(trade_day,code):
                    pass
                else:
                    continue
                
                if (
                        (self.stockclose_sht.loc[trade_day,code]
                         *self.adjfac_sht.loc[trade_day,code])
                            >= 
                        self.calcu_MA(trade_day,code,period=30)
                    ):
                    selected_index.append(j)
                if len(selected_index)>=self.num_of_stock:
                    break
            tmp1=copy.deepcopy(love_sheet.loc[selected_index,:].iloc[:,:2])
            my_portfolio1[trade_day]=copy.deepcopy(tmp1.reset_index(drop=True))
        self.my_portfolio=copy.deepcopy(my_portfolio1)
        
        
class MyPortfolio(MyPortfolio2):
    def __init__(self,datadict,num_of_stock):
        super().__init__(datadict,num_of_stock)
        
    def get_para_data(self):
        self.calcu_daily_holdings()
        output={"adjdatelist":self.adjdatelist,
                "adjfac_sht":self.adjfac_sht,
                "stockclose_sht":self.stockclose_sht,
                "datelist":self.datelist,
                "fundlove":self.fundlove,
                "num_of_stock":self.num_of_stock
            }
        return copy.deepcopy(output)
    
    def get_pf(self):
        return self.my_portfolio
    
        
class CalcuNav():
    def __init__(self,my_portfolio,basic_parameter):
        self.datelist=basic_parameter["datelist"]
        self.fundlove=basic_parameter["fundlove"]
        self.num_of_stock=basic_parameter["num_of_stock"]
        self.stockclose_sht=basic_parameter["stockclose_sht"]
        self.adjfac_sht=basic_parameter["adjfac_sht"]
        self.adjdatelist=basic_parameter["adjdatelist"]
        self.my_portfolio1=copy.deepcopy(my_portfolio)

    def calcu_each_value(self,i,hold,hold_ex):
        dex1=self.datelist[i-1]
        d1=self.datelist[i]
        hold1=copy.deepcopy(hold)
        for i1 in hold1.index:
            code1=hold1.loc[i1,"代码"]
            hold1.loc[i1,"价值"]=(
                                (self.stockclose_sht.loc[d1,code1]*self.adjfac_sht.loc[d1,code1]) /
                                (self.stockclose_sht.loc[dex1,code1]*self.adjfac_sht.loc[dex1,code1])
                                    )*hold_ex.loc[i1,"价值"]
        return hold1

    def add_close_and_fac(self,hold,on):
        faclist=[]
        closelist=[]
        for i in range(len(hold.index)):
            the_day=hold[on].iat[i]
            the_code=hold["代码"].iat[i]
            tmp1=self.stockclose_sht.loc[the_day,the_code]
            tmp2=self.adjfac_sht.loc[the_day,the_code]
            closelist.append(tmp1)
            faclist.append(tmp2)
        if on=="上一调仓日":
            hold["调仓日复权因子"]=faclist
            hold["调仓日收盘价"]=closelist
        else:
            hold["当日复权因子"]=faclist
            hold["当日收盘价"]=closelist
        return hold

    def  adjust_stock_but_value_nochange(self,hold,hold_middle_ex):
        h1=copy.deepcopy(hold)
        h1["价值"]=0.0
        h2=copy.deepcopy(hold_middle_ex)
        rest=0.0
        n1=0
        for i1 in range(len(h2.index)):
            if h2.loc[i1,"代码"] in h1["代码"].values:
                pass
            else:
                rest=rest+h2.loc[i1,"价值"]
                n1=n1+1
        for i1 in range(len(h1.index)):
            if h1.loc[i1,"代码"] in h2["代码"].values:
                h1.loc[i1,"价值"]=h2[h2["代码"]==h1.loc[i1,"代码"]]["价值"].iat[0]
            else:
                h1.loc[i1,"价值"]=rest/n1
        return h1


    def do_and_return(self):
        for i in range(len(self.datelist)):
        #for i in range(116): 
            #i=134,74,116
            trade_day=self.datelist[i]
            if trade_day<self.adjdatelist["adjday"][0]:
                continue
            print(i/len(self.datelist))
            
            if trade_day==self.adjdatelist["adjday"][0]:
                hold=copy.deepcopy(self.my_portfolio1[trade_day])
                hold["当日"]=[trade_day for i in range(len(hold.index))]
                hold=self.add_close_and_fac(hold,on="当日")
                hold["价值"]=[1/self.num_of_stock for i in range(len(hold.index))]
                self.my_portfolio1[trade_day]=hold
            else:
                hold_ex=copy.deepcopy(self.my_portfolio1[self.datelist[i-1]])
                hold=copy.deepcopy(hold_ex)
                hold_middle_ex=copy.deepcopy(self.calcu_each_value(i,hold,hold_ex))
                hold_middle_ex["价值"].sum()
                hold=copy.deepcopy(self.my_portfolio1[self.datelist[i]])
                hold["当日"]=[trade_day for i in range(len(hold.index))]
                hold=self.add_close_and_fac(hold,on="当日")
                if set(hold["名称"])==set(hold_middle_ex["名称"]):
                    hold=copy.deepcopy(hold_middle_ex)
                else:
                    hold=self.adjust_stock_but_value_nochange(hold,hold_middle_ex)
                self.my_portfolio1[trade_day]=hold
        return self.my_portfolio1

class MyPortfolioDueFundsNum(MyPortfolio):
    def __init__(self,datadict,num_of_stock):
        super().__init__(datadict,num_of_stock)
        
    def init_price_and_factor(self):
        for i1 in self.stockclose_dict.values():
            for stockclose_sht in i1.values():
                break
        self.stockclose_sht=stockclose_sht.copy()
        self.stockclose_sht.index=stockclose_sht["date"]
        del self.stockclose_sht["date"]
        for i1 in self.adjfactor_dict.values():
            for adjfac_sht in i1.values():
                break
        self.adjfac_sht=adjfac_sht.copy()
        self.adjfac_sht.index=adjfac_sht["date"]
        
        for i,j in self.fundlove.items():
            tmp1=list(j.values())[0]
            tmp1=tmp1.sort_values(by="持有基金数",ascending=False).reset_index(drop=True)
            self.fundlove[i]=copy.deepcopy({"file":tmp1})

        
        
class MyPortfolioDueAssetRatio(MyPortfolio):
    def __init__(self,datadict,num_of_stock):
        super().__init__(datadict,num_of_stock)
        
    def init_price_and_factor(self):
        for i1 in self.stockclose_dict.values():
            for stockclose_sht in i1.values():
                break
        self.stockclose_sht=stockclose_sht.copy()
        self.stockclose_sht.index=stockclose_sht["date"]
        del self.stockclose_sht["date"]
        for i1 in self.adjfactor_dict.values():
            for adjfac_sht in i1.values():
                break
        self.adjfac_sht=adjfac_sht.copy()
        self.adjfac_sht.index=adjfac_sht["date"]
        
        for i,j in self.fundlove.items():
            tmp1=list(j.values())[0]
            #tmp1["ratio"]=tmp1["持仓市值"]/tmp1["流通市值"]
            #tmp1["ratio"]=tmp1["ratio"].map(lambda x:0.0 if x==np.inf else x)
            tmp1=tmp1.sort_values(by="占流通股比",ascending=False,na_position="last").reset_index(drop=True)
            self.fundlove[i]=copy.deepcopy({"file":tmp1})        
        

class MyPortfolioDueFundsChange(MyPortfolio):
    def __init__(self,datadict,num_of_stock):
        super().__init__(datadict,num_of_stock)
        
    def init_price_and_factor(self):
        for i1 in self.stockclose_dict.values():
            for stockclose_sht in i1.values():
                break
        self.stockclose_sht=stockclose_sht.copy()
        self.stockclose_sht.index=stockclose_sht["date"]
        del self.stockclose_sht["date"]
        for i1 in self.adjfactor_dict.values():
            for adjfac_sht in i1.values():
                break
        self.adjfac_sht=adjfac_sht.copy()
        self.adjfac_sht.index=adjfac_sht["date"]
        
        for i,j in self.fundlove.items():
            tmp1=list(j.values())[0]
            tmp1=tmp1.sort_values(by="基金增减数量",ascending=False,na_position="last").reset_index(drop=True)
            self.fundlove[i]=copy.deepcopy({"file":tmp1})        

class MyPortfolioDueFundsChangeNoMA(MyPortfolioDueFundsChange):
    def __init__(self,datadict,num_of_stock):
        super().__init__(datadict,num_of_stock)
        
    def calcu_daily_holdings(self):
        self.init_price_and_factor()
        self.create_adjustdatelist()
        del self.stockclose_dict,self.adjfactor_dict
        
        my_portfolio1=dict()
        for i in range(len(self.datelist)):
        #for i in range(115):
            #i=74,115
            trade_day=self.datelist[i]
            if trade_day<self.adjdatelist["adjday"][0]:
                continue
            print(i/len(self.datelist))            
            j=0
            while True:
                if (trade_day<self.adjdatelist["adjday"][j+1]) and (self.adjdatelist["adjday"][j]<=trade_day):
                    break
                else:
                    j=j+1
                if j>=len(self.adjdatelist.index)-1:
                    break
            filename=self.adjdatelist["filename"][j]
            love_sheet=list(self.fundlove[filename].values())[0]
            selected_index=list()
            
            for j in range(len(love_sheet.index)):
                code=love_sheet.loc[j,"代码"]
                if (code in self.stockclose_sht.columns) and (code in self.adjfac_sht.columns):
                    pass
                else:
                    continue
                if self.is_valid(trade_day,code):
                    pass
                else:
                    continue
                
                if True:
                    selected_index.append(j)
                if len(selected_index)>=self.num_of_stock:
                    break
            tmp1=copy.deepcopy(love_sheet.loc[selected_index,:].iloc[:,:2])
            my_portfolio1[trade_day]=copy.deepcopy(tmp1.reset_index(drop=True))
        self.my_portfolio=copy.deepcopy(my_portfolio1)
   
    
        
        
        
class MyPortfolioDueAssetRatioNoMA(MyPortfolioDueAssetRatio):
    def __init__(self,datadict,num_of_stock):
        super().__init__(datadict,num_of_stock)
        
    def calcu_daily_holdings(self):
        self.init_price_and_factor()
        self.create_adjustdatelist()
        del self.stockclose_dict,self.adjfactor_dict
        
        my_portfolio1=dict()
        for i in range(len(self.datelist)):
        #for i in range(115):
            #i=74,115
            trade_day=self.datelist[i]
            if trade_day<self.adjdatelist["adjday"][0]:
                continue
            print(i/len(self.datelist))            
            j=0
            while True:
                if (trade_day<self.adjdatelist["adjday"][j+1]) and (self.adjdatelist["adjday"][j]<=trade_day):
                    break
                else:
                    j=j+1
                if j>=len(self.adjdatelist.index)-1:
                    break
            filename=self.adjdatelist["filename"][j]
            love_sheet=list(self.fundlove[filename].values())[0]
            selected_index=list()
            
            for j in range(len(love_sheet.index)):
                code=love_sheet.loc[j,"代码"]
                if (code in self.stockclose_sht.columns) and (code in self.adjfac_sht.columns):
                    pass
                else:
                    continue
                if self.is_valid(trade_day,code):
                    pass
                else:
                    continue
                
                if True:
                    selected_index.append(j)
                if len(selected_index)>=self.num_of_stock:
                    break
            tmp1=copy.deepcopy(love_sheet.loc[selected_index,:].iloc[:,:2])
            my_portfolio1[trade_day]=copy.deepcopy(tmp1.reset_index(drop=True))
        self.my_portfolio=copy.deepcopy(my_portfolio1)
        
        
        
class MyPortfolioDueFundsNumNoMA(MyPortfolioDueFundsNum):
    def __init__(self,datadict,num_of_stock):
        super().__init__(datadict,num_of_stock)
        
    def calcu_daily_holdings(self):
        self.init_price_and_factor()
        self.create_adjustdatelist()
        del self.stockclose_dict,self.adjfactor_dict
        
        my_portfolio1=dict()
        for i in range(len(self.datelist)):
        #for i in range(115):
            #i=74,115
            trade_day=self.datelist[i]
            if trade_day<self.adjdatelist["adjday"][0]:
                continue
            print(i/len(self.datelist))            
            j=0
            while True:
                if (trade_day<self.adjdatelist["adjday"][j+1]) and (self.adjdatelist["adjday"][j]<=trade_day):
                    break
                else:
                    j=j+1
                if j>=len(self.adjdatelist.index)-1:
                    break
            filename=self.adjdatelist["filename"][j]
            love_sheet=list(self.fundlove[filename].values())[0]
            selected_index=list()
            
            for j in range(len(love_sheet.index)):
                code=love_sheet.loc[j,"代码"]
                if (code in self.stockclose_sht.columns) and (code in self.adjfac_sht.columns):
                    pass
                else:
                    continue
                if self.is_valid(trade_day,code):
                    pass
                else:
                    continue
                
                if True:
                    selected_index.append(j)
                if len(selected_index)>=self.num_of_stock:
                    break
            tmp1=copy.deepcopy(love_sheet.loc[selected_index,:].iloc[:,:2])
            my_portfolio1[trade_day]=copy.deepcopy(tmp1.reset_index(drop=True))
        self.my_portfolio=copy.deepcopy(my_portfolio1)
        
class MyPortfolioNoMA(MyPortfolio):
    def __init__(self,datadict,num_of_stock):
        super().__init__(datadict,num_of_stock)
        
    def calcu_daily_holdings(self):
        self.init_price_and_factor()
        self.create_adjustdatelist()
        del self.stockclose_dict,self.adjfactor_dict
        
        my_portfolio1=dict()
        for i in range(len(self.datelist)):
        #for i in range(115):
            #i=74,115
            trade_day=self.datelist[i]
            if trade_day<self.adjdatelist["adjday"][0]:
                continue
            print(i/len(self.datelist))            
            j=0
            while True:
                if (trade_day<self.adjdatelist["adjday"][j+1]) and (self.adjdatelist["adjday"][j]<=trade_day):
                    break
                else:
                    j=j+1
                if j>=len(self.adjdatelist.index)-1:
                    break
            filename=self.adjdatelist["filename"][j]
            love_sheet=list(self.fundlove[filename].values())[0]
            selected_index=list()
            
            for j in range(len(love_sheet.index)):
                code=love_sheet.loc[j,"代码"]
                if (code in self.stockclose_sht.columns) and (code in self.adjfac_sht.columns):
                    pass
                else:
                    continue
                if self.is_valid(trade_day,code):
                    pass
                else:
                    continue
                
                if True:
                    selected_index.append(j)
                if len(selected_index)>=self.num_of_stock:
                    break
            tmp1=copy.deepcopy(love_sheet.loc[selected_index,:].iloc[:,:2])
            my_portfolio1[trade_day]=copy.deepcopy(tmp1.reset_index(drop=True))
        self.my_portfolio=copy.deepcopy(my_portfolio1)
        
        
        
        
        
        
