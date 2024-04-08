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

def read_one_after_one(dirpath,date=True):
        dirlist=os.listdir(dirpath)
        greatlis=dict()
        filenamelis=list()
        for filename in dirlist:
            print(filename)
            filenamelis.append(filename)
            if os.path.isdir(os.path.join(dirpath,filename)) or filename=="ok.xlsx" or ("~$" in filename):
                continue
            wb=app.books.open(os.path.join(dirpath,filename))
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
    
def just_pile(greatlis,ept):
    sheetset=set()
    pile_up_dict=dict()
    ept_dict=dict()
    greatlis=copy.deepcopy(greatlis)
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

def show_side_number(greatlis,axis="both"):
    greatlis.keys()
    for filename in greatlis.keys():
        filename
        for shtname in greatlis[filename].keys():
            shtname
            greatlis[filename][shtname]
            print(filename,shtname,len(greatlis[filename][shtname].columns),
                  len(greatlis[filename][shtname].index))
           
def get_datelist(dictlist,level=1):
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
    
def fundlove_get_top(sheetdict,topnumber=10,use_hk=True):
    for i in sheetdict.keys():
        sheet=sheetdict[i]
    if use_hk:
        top=sheet.iloc[:topnumber,:2].copy()
    else:
        f=0
        top=pd.DataFrame(columns=["代码","名称"])
        while True:
            if f>=len(sheet.index) or len(top.index)==topnumber:
                break
            if sheet["代码"].iat[f][-2:]=="HK":
                f=f+1
                continue
            else:
                tmp1=pd.DataFrame([sheet["代码"].iat[f],sheet["名称"].iat[f]],index=["代码","名称"]).T
                top=pd.concat([top,tmp1],axis=0)
                f=f+1
        top=top.reset_index(drop=True)
    return top

def search_adjust_day(adjust_day):
    i=0
    while True:
        if adjust_day+datetime.timedelta(days=i) in datelist.values:
            break
        else:
            i=i+1
    adjust_day=adjust_day+datetime.timedelta(days=i)
    adjindex=datelist[datelist==adjust_day].index[0]+15
    adjust_day=datelist.loc[adjindex]
    return adjust_day

def add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on):
    for i in stockclose_dict.values():
        for stockclose_sht in i.values():
            break
    stockclose_sht=stockclose_sht.copy()
    stockclose_sht.index=stockclose_sht["date"]
    del stockclose_sht["date"]
    for i in adjfactor_dict.values():
        for adjfac_sht in i.values():
            break
    adjfac_sht=adjfac_sht.copy()
    adjfac_sht.index=adjfac_sht["date"]
    faclist=[]
    closelist=[]
    for i in range(len(hold.index)):
        the_day=hold[on].iat[i]
        the_code=hold["代码"].iat[i]
        tmp1=stockclose_sht.loc[the_day,the_code]
        tmp2=adjfac_sht.loc[the_day,the_code]
        closelist.append(tmp1)
        faclist.append(tmp2)
    if on=="上一调仓日":
        hold["调仓日复权因子"]=faclist
        hold["调仓日收盘价"]=closelist
    else:
        hold["当日复权因子"]=faclist
        hold["当日收盘价"]=closelist
    return hold

def change_money_pct(hold,hold_ex_to_now,blacklist):
    money=hold_ex_to_now["价值"].sum()
    hold["价值"]=[0.0 for j in range(len(hold.index))]
    print("调出",set(hold_ex_to_now["名称"]) - set(hold["名称"]))
    print("调入",set(hold["名称"]) - set(hold_ex_to_now["名称"]))
    rest_money=0
    for i in range(len(hold_ex_to_now.index)):
        tmp1=hold_ex_to_now["代码"].iat[i]
        tmp2=hold_ex_to_now["价值"].iat[i]
        if tmp1 in hold["代码"].values:
            idx=hold[hold["代码"]==tmp1].index[0]
            if tmp2/money>0.1:
                hold.loc[idx,"价值"]=money*0.1
                rest_money=rest_money+tmp2-money*0.1
            elif (tmp1 in blacklist.keys()) and blacklist[tmp1]["status"]=="on_jail":
                hold.loc[idx,"价值"]=0
                rest_money=rest_money+tmp2
                print(i)
            else:
                hold.loc[idx,"价值"]=tmp2
        else:
            rest_money=rest_money+tmp2
    tmp1=0
    for i in range(len(hold.index)):
        tmp2=hold["价值"].iat[i]
        tmp3=hold["代码"].iat[i]
        if tmp2==0 and not(tmp3 in blacklist.keys()):
            tmp1=tmp1+1
    for i in range(len(hold.index)):
        tmp2=hold["价值"].iat[i]
        tmp3=hold["代码"].iat[i]
        if tmp2==0 and not(tmp3 in blacklist.keys()):
            hold["价值"].iat[i]=rest_money/tmp1
    return money,hold

def check_blacklist(hold,stockclose_dict,adjfactor_dict,the_day,blacklist):
    def get_rid_of_skin(dict_all):
        for i in dict_all.values():
            for sht in i.values():
                break
        sht=sht.copy()
        sht.index=sht["date"]
        del sht["date"]
        return sht
    def get_start_date():
        st=(the_day-pd.to_datetime("2018-01-01")).days
        st=0 if st-120<=0 else st-120
        while True:
            if st==0:
                break
            if (the_day-pd.Timedelta(days=st)) in get_datelist([adjfactor_dict,stockclose_dict]).values:
                break
            else:
                st=st-1
        st=pd.to_datetime("2018-01-01")+pd.Timedelta(days=st)
        return st
    def divide_to_others_evenly(hold):
        rest_money=hold["价值"].iat[i].copy()
        hold["价值"].iat[i]=0
        tmp4=0
        for k in range(len(hold.index)):
            if k!=i and not(hold["代码"].iat[k] in blacklist.keys() and (blacklist[hold["代码"].iat[k]]["status"]=="on_jail")):
                tmp4=tmp4+1
        for k in range(len(hold.index)):
            if k==i or (hold["代码"].iat[k] in blacklist.keys()) and (blacklist[hold["代码"].iat[k]]["status"]=="on_jail"):
                continue 
            else:
                hold["价值"].iat[k]=hold["价值"].iat[k]+rest_money/tmp4
        return hold
    def give_nav_from_all_others(hold):
        tmp4=0
        for k in range(len(hold.index)):
            if k!=i and hold["价值"].iat[k]>0:
                tmp4=tmp4+1
        rest_money=hold["价值"].sum()/(tmp4+1)
        for k in range(len(hold.index)):
            if k==i:
                hold["价值"].iat[k]=rest_money
            elif hold["价值"].iat[k]>0:
                hold["价值"].iat[k]=hold["价值"].iat[k]*tmp4/(tmp4+1)
        return hold

    stockclose_sht=get_rid_of_skin(stockclose_dict)
    adjfac_sht=get_rid_of_skin(adjfactor_dict)
    for i in range(len(hold.index)):
        the_code=hold["代码"].iat[i]
        tmp1=stockclose_sht.loc[the_day,the_code]
        tmp2=adjfac_sht.loc[the_day,the_code]
        st=get_start_date()
        tmp3=stockclose_sht.loc[st:the_day,the_code]*adjfac_sht.loc[st:the_day,the_code]
        if the_code in blacklist.keys():
            if blacklist[the_code]["status"]=="have_type_A_loss_stop":
                if (tmp1*tmp2)>tmp3.min()*1.5:
                    print("it goes to the edge type B!",i,the_day)
                    blacklist[the_code]["type_B_warning_price"]= (tmp1*tmp2)*0.8
                    blacklist[the_code]["status"]="have_type_AB_loss_stop"
            elif blacklist[the_code]["status"]=="have_type_AB_loss_stop":
                if tmp1*tmp2>(blacklist[the_code]["type_B_warning_price"]/0.8):
                    print("the type B edge changes!",i,the_day)
                    blacklist[the_code]["type_B_warning_price"]=tmp1*tmp2*0.8
                    blacklist[the_code]["status"]="have_type_AB_loss_stop"
        else:
            blacklist[the_code]={}
            blacklist[the_code]["status"]="have_type_A_loss_stop"
            blacklist[the_code]["type_A_warning_price"]=(tmp1*tmp2)*0.8
            
        if the_code in blacklist.keys():  
            if blacklist[the_code]["status"]=="have_type_A_loss_stop":
                if (tmp1*tmp2)<=blacklist[the_code]["type_A_warning_price"]:
                    print("it goes to the jail because of A!",i,the_day)
                    blacklist[the_code]["status"]="on_jail"
                    blacklist[the_code]["out_when_price_is"]=blacklist[the_code]["type_A_warning_price"]/0.8
                    blacklist[the_code].pop("type_A_warning_price")
                    hold=divide_to_others_evenly(hold)
                    hold["价值"].sum()
            elif blacklist[the_code]["status"]=="have_type_AB_loss_stop":
                if (tmp1*tmp2)<=blacklist[the_code]["type_A_warning_price"]:
                    print("it goes to the jail because of A!",i,the_day)
                    blacklist[the_code]["status"]="on_jail"
                    blacklist[the_code]["out_when_price_is"]=blacklist[the_code]["type_A_warning_price"]/0.8
                    blacklist[the_code].pop("type_A_warning_price")
                    blacklist[the_code].pop("type_B_warning_price")
                    hold=divide_to_others_evenly(hold)
                    hold["价值"].sum()
                elif (tmp1*tmp2)<=blacklist[the_code]["type_B_warning_price"]:
                    print("it goes to the jail because of B!",i,the_day)
                    blacklist[the_code]["status"]="on_jail"
                    blacklist[the_code]["out_when_price_is"]=blacklist[the_code]["type_B_warning_price"]/0.8
                    blacklist[the_code].pop("type_A_warning_price")
                    blacklist[the_code].pop("type_B_warning_price")
                    hold=divide_to_others_evenly(hold)
    
            if blacklist[the_code]["status"]=="on_jail":
                if (tmp1*tmp2)>=blacklist[the_code]["out_when_price_is"]:
                    print("it goes out because of hero's back!!",i,the_day)
                    blacklist.pop(the_code)
                    hold=give_nav_from_all_others(hold)
                    hold["价值"].sum()
                elif (tmp1*tmp2)<=blacklist[the_code]["out_when_price_is"]*0.5:
                    print("it goes out because it's too low!!",i,the_day)
                    blacklist.pop(the_code)
                    hold=give_nav_from_all_others(hold)
                    hold["价值"].sum()                    
    return hold,blacklist

def adjust_according_to_price():    
    flagdatelist=pd.Series(list(my_portfolio.keys())).sort_values(ascending=True).reset_index(drop=True)
    for i in range(len(datelist)):
        #i=74
        trade_day=datelist[i]
        if trade_day<flagdatelist[0]:
            continue
        print(i/len(datelist))
        if trade_day in flagdatelist.values:
            hold=my_portfolio[trade_day]["holdings"].copy()
            hold["上一调仓日"]=[trade_day for i in range(len(hold.index))] 
            hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="上一调仓日")
            hold["当日复权因子"]=hold["调仓日复权因子"].copy()
            hold["当日收盘价"]=hold["调仓日收盘价"].copy()
            my_portfolio[trade_day]["holdings"]=hold       
        else:
            my_portfolio[trade_day]=dict()
            hold_ex=my_portfolio[datelist[i-1]]["holdings"].copy()
            hold=hold_ex.copy()
            del hold["当日复权因子"]
            del hold["当日收盘价"]
            hold["当日"]=[trade_day for i in range(len(hold.index))]
            hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="当日")
            del hold["当日"]
            my_portfolio[trade_day]["holdings"]=hold.copy()
    
    datelist_new=pd.Series(list(my_portfolio.keys())).sort_values(ascending=True).reset_index(drop=True)
    blacklist={}
    for i in range(len(datelist_new)):
    #for i in range(497):
        trade_day=datelist_new[i]
        if i==0:
            money=1
            my_portfolio[trade_day]["money"]=money
            continue
        hold=my_portfolio[trade_day]["holdings"].copy()
        trade_day_ex=datelist_new[i-1]
        hold_ex=my_portfolio[trade_day_ex]["holdings"].copy()
        if trade_day in hold["上一调仓日"].values:
            hold_ex_to_now=hold_ex.copy()
            del hold_ex_to_now["当日复权因子"]
            del hold_ex_to_now["当日收盘价"]
            hold_ex_to_now["当日"]=[trade_day for i in range(len(hold_ex_to_now.index))]
            hold_ex_to_now=add_close_and_fac(hold_ex_to_now,stockclose_dict,adjfactor_dict,on="当日")
            del hold_ex_to_now["当日"]
            for j in range(len(hold.index)):
                ctmp1=hold_ex_to_now["当日收盘价"].iat[j]
                ftmp1=hold_ex_to_now["当日复权因子"].iat[j]
                tmp1=hold_ex_to_now["代码"].iat[j]
                k=hold_ex[hold_ex["代码"]==tmp1].index[0]
                ctmp2=hold_ex.loc[k,"当日复权因子"]
                ftmp2=hold_ex.loc[k,"当日收盘价"]
                hold_ex_to_now["价值"].iat[j]=hold_ex.loc[k,"价值"]*ctmp1*ftmp1/(ctmp2*ftmp2)
            money,hold=change_money_pct(hold,hold_ex_to_now,blacklist)
            my_portfolio[trade_day]["money"]=money
            my_portfolio[trade_day]["holdings"]=hold
        else:
            for j in range(len(hold.index)):
                ctmp1=hold["当日收盘价"].iat[j]
                ftmp1=hold["当日复权因子"].iat[j]
                tmp1=hold["代码"].iat[j]
                k=hold_ex[hold_ex["代码"]==tmp1].index[0]
                ctmp2=hold_ex.loc[k,"当日复权因子"]
                ftmp2=hold_ex.loc[k,"当日收盘价"]
                hold["价值"].iat[j]=hold_ex.loc[k,"价值"]*ctmp1*ftmp1/(ctmp2*ftmp2)
            hold,blacklist=check_blacklist(hold,stockclose_dict,adjfactor_dict,trade_day,blacklist)
            money=hold["价值"].sum()
            my_portfolio[trade_day]["money"]=money
            my_portfolio[trade_day]["holdings"]=hold
    return my_portfolio

def adjust_according_to_loverank():
    def remove_dec_stock():
        my_portfolio1=my_portfolio.copy()
        for i in range(len(adjust_date_list)):
            today=adjust_date_list[i]
            if i==0:
                continue
            ex_day=adjust_date_list[i-1]
            h=my_portfolio1[today]["holdings"].copy()
            h_ex=my_portfolio1[ex_day]["holdings"].copy()
            del h["价值"],h_ex["价值"]
            h["date"]=today
            h_ex["date"]=ex_day
            #h_all=pd.merge(h,h_ex,how="outer",on=["代码","名称"],suffixes=("today","ex_day"))
            inc=[]
            dec=[]
            for code in (set(h["代码"]) | set(h_ex["代码"])):
                if (code in h["代码"].values) and (code in h_ex["代码"].values):
                    idx1=h[h["代码"]==code].index[0]
                    idx2=h_ex[h_ex["代码"]==code].index[0]
                    if idx1<=idx2:
                        inc.append(code)
                    else:
                        dec.append(code)
                elif (code in h["代码"].values):
                    inc.append(code)
                elif (code in h_ex["代码"].values):
                    dec.append(code)
            for j in range(len(h["代码"].index)):
                if (h.loc[j,"代码"] in inc):
                    pass
                else:
                    h=h.drop(j,axis=0)
            
            h=h.reset_index(drop=True).drop("date",axis=1)
            h["价值"]=1/len(h.index)
            my_portfolio1[today]["holdings"]=h.copy()
        return my_portfolio1
    
    def calcu_each_value():
        dex1=datelist[i-1]
        d1=datelist[i]
        hold1=hold.copy()
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
        for i1 in hold1.index:
            hold1.loc[i1,"价值"]
            hold_ex.loc[i1,"价值"]
            code1=hold1.loc[i1,"代码"]
            hold1.loc[i1,"价值"]=(
                                (stockclose_sht.loc[d1,code1]*adjfac_sht.loc[d1,code1]) /
                                (stockclose_sht.loc[dex1,code1]*adjfac_sht.loc[dex1,code1])
                                    )*hold_ex.loc[i1,"价值"]
        return hold1

    def do_adjust():
        hold1=copy.deepcopy(hold)
        holdex1=copy.deepcopy(hold_middle_ex)
        dex1=datelist[i-1]
        d1=datelist[i]
        holdex1["价值"].sum()
        hold1["价值"].sum()
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
        for i1 in hold1.index:
            code1=hold1.loc[i1,"代码"]
            if code1 in holdex1["代码"].values:
                hold1.loc[i1,"价值"]=holdex1[holdex1["代码"]==code1]["价值"].iat[0]
            else:
                hold1.loc[i1,"价值"]=0.0
        num1=len(hold1[hold1["价值"]==0.0].index)
        num1=(holdex1["价值"].sum()-hold1["价值"].sum())/num1
        for i1 in hold1.index:
            if hold1.loc[i1,"价值"]==0.0:
                hold1.loc[i1,"价值"]=num1
        return hold1

    adjust_date_list=pd.Series(
        list(my_portfolio.keys())).sort_values(ascending=True
                                               ).reset_index(drop=True)
    my_portfolio1=copy.deepcopy(remove_dec_stock())
    for i in range(len(datelist)):
    #for i in range(134): 
        #i=134,74
        trade_day=datelist[i]
        if trade_day<adjust_date_list[0]:
            continue
        print(i/len(datelist))
        if trade_day in adjust_date_list.values:
            if trade_day==adjust_date_list[0]:
                hold=my_portfolio1[trade_day]["holdings"].copy()
                hold["上一调仓日"]=[trade_day for i in range(len(hold.index))] 
                hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="上一调仓日")
                hold["当日复权因子"]=hold["调仓日复权因子"].copy()
                hold["当日收盘价"]=hold["调仓日收盘价"].copy()
                my_portfolio1[trade_day]["holdings"]=hold
            else:
                hold_ex=copy.deepcopy(my_portfolio1[datelist[i-1]]["holdings"])
                hold=copy.deepcopy(hold_ex)
                hold_middle_ex=copy.deepcopy(calcu_each_value())
                hold=copy.deepcopy(my_portfolio1[trade_day]["holdings"])
                hold["上一调仓日"]=[trade_day for i in range(len(hold.index))]
                hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="上一调仓日")
                hold["当日复权因子"]=hold["调仓日复权因子"].copy()
                hold["当日收盘价"]=hold["调仓日收盘价"].copy()
                hold["价值"].sum()
                hold=do_adjust()
                my_portfolio1[trade_day]["holdings"]=hold
        else:
            my_portfolio1[trade_day]=dict()
            hold_ex=my_portfolio1[datelist[i-1]]["holdings"].copy()
            hold=hold_ex.copy()
            del hold["当日复权因子"]
            del hold["当日收盘价"]
            hold["当日"]=[trade_day for i in range(len(hold.index))]
            hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="当日")
            del hold["当日"]
            hold=calcu_each_value().copy()
            my_portfolio1[trade_day]["holdings"]=hold.copy()
    return my_portfolio1
        
def adjust_according_to_MA():
    def init_price_and_factor():
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
        return stockclose_sht,adjfac_sht
    def calcu_each_value():
        dex1=datelist[i-1]
        d1=datelist[i]
        hold1=copy.deepcopy(hold)
        for i1 in hold1.index:
            hold1.loc[i1,"价值"]
            hold_ex.loc[i1,"价值"]
            code1=hold1.loc[i1,"代码"]
            hold1.loc[i1,"价值"]=(
                                (stockclose_sht.loc[d1,code1]*adjfac_sht.loc[d1,code1]) /
                                (stockclose_sht.loc[dex1,code1]*adjfac_sht.loc[dex1,code1])
                                    )*hold_ex.loc[i1,"价值"]
        return hold1
    def check_blk():
        def calcu_MA(period=30):
            if datelist[datelist==trade_day].index[0]<=period:
                idx1=0
            else:
                idx1=datelist[datelist==trade_day].index[0]-period
            dtlis=datelist.iloc[idx1:idx1+period]
            MA1=(stockclose_sht.loc[dtlis,the_code]*adjfac_sht.loc[dtlis,the_code]).mean()
            return MA1
        def gojail():
            hold2=copy.deepcopy(hold1)
            rest_money2=hold2["价值"].iat[i1]
            hold2["价值"].iat[i1]=0
            t2=0
            for k in range(len(hold2.index)):
                if k!=i1 and not(hold2["代码"].iat[k] in blacklist1.keys() and (blacklist1[hold2["代码"].iat[k]]["status"]=="on_jail")):
                    t2=t2+1
            for k in range(len(hold2.index)):
                if k==i1 or (hold2["代码"].iat[k] in blacklist1.keys()) and (blacklist1[hold2["代码"].iat[k]]["status"]=="on_jail"):
                    continue 
                else:
                    hold2["价值"].iat[k]=hold2["价值"].iat[k]+rest_money2/t2
            hold2["价值"].sum()
            if t2==0:
                return hold2,rest_money2
            else:
                return hold2,0.0
        def outjail():
            hold2=copy.deepcopy(hold1)
            t2=0
            for k in range(len(hold2.index)):
                if k!=i1 and hold2["价值"].iat[k]>0:
                    t2=t2+1
            if t2==0:
                rest_money2=(hold2["价值"].sum()+rest1)/(t2+1)
            else:
                rest_money2=(hold2["价值"].sum())/(t2+1)
            for k in range(len(hold2.index)):
                if k==i1:
                    hold2["价值"].iat[k]=rest_money2
                elif hold2["价值"].iat[k]>0:
                    hold2["价值"].iat[k]=hold2["价值"].iat[k]*t2/(t2+1)
            return hold2,0.0
        hold1=copy.deepcopy(hold)
        rest1=rest
        blacklist1=copy.deepcopy(blacklist)
        for i1 in range(len(hold.index)):
            the_code=hold["代码"].iat[i1]
            tmp1=stockclose_sht.loc[trade_day,the_code]
            tmp2=adjfac_sht.loc[trade_day,the_code]
            the_MA=calcu_MA(30)
            if the_code in blacklist1.keys():
                if tmp1*tmp2>=the_MA:
                    print("hero back!",i1,hold1["名称"].iat[i1],trade_day)
                    blacklist1.pop(the_code)
                    hold1,rest1=outjail()
            else:
                #不在黑名单
                if tmp1*tmp2<=the_MA:
                    print("it goes to jail!",i1,hold1["名称"].iat[i1],trade_day)
                    blacklist1[the_code]={}
                    blacklist1[the_code]["status"]="on_jail"
                    hold1,rest1=gojail()
        return hold1,blacklist1,rest1
    def do_adjust():
        h1=copy.deepcopy(hold)
        hex1=copy.deepcopy(hold_middle_ex)
        sum1=hex1["价值"].sum()
        for i1 in h1.index:
            h1.loc[i1,"价值"]=sum1/len(h1.index)
        return h1
        
        
        
        
    stockclose_sht,adjfac_sht=init_price_and_factor()
    adjust_date_list=pd.Series(
        list(my_portfolio.keys())).sort_values(ascending=True
                                               ).reset_index(drop=True)
    my_portfolio1=copy.deepcopy(my_portfolio)
    blacklist,rest={},0
    
    for i in range(len(datelist)):
    #for i in range(150): 
        #i=134,74,151
        trade_day=datelist[i]
        if trade_day<adjust_date_list[0]:
            continue
        print(i/len(datelist))
        if trade_day in adjust_date_list.values:
            if trade_day==adjust_date_list[0]:
                hold=copy.deepcopy(my_portfolio1[trade_day]["holdings"])
                hold["上一调仓日"]=[trade_day for i in range(len(hold.index))] 
                hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="上一调仓日")
                hold["当日复权因子"]=hold["调仓日复权因子"].copy()
                hold["当日收盘价"]=hold["调仓日收盘价"].copy()
                my_portfolio1[trade_day]["holdings"]=hold
                my_portfolio1[trade_day]["money"]=hold["价值"].sum()
            else:
                hold_ex=copy.deepcopy(my_portfolio1[datelist[i-1]]["holdings"])
                hold=copy.deepcopy(hold_ex)
                hold_middle_ex=copy.deepcopy(calcu_each_value())
                hold=copy.deepcopy(my_portfolio1[trade_day]["holdings"])
                hold["上一调仓日"]=[trade_day for i in range(len(hold.index))]
                hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="上一调仓日")
                hold["当日复权因子"]=hold["调仓日复权因子"].copy()
                hold["当日收盘价"]=hold["调仓日收盘价"].copy()
                hold["价值"].sum()
                hold=do_adjust()
                blacklist=dict()
                my_portfolio1[trade_day]["holdings"]=hold
                my_portfolio1[trade_day]["money"]=hold["价值"].sum()
        else:
            my_portfolio1[trade_day]=dict()
            hold_ex=copy.deepcopy(my_portfolio1[datelist[i-1]]["holdings"])
            hold=copy.deepcopy(hold_ex)
            del hold["当日复权因子"]
            del hold["当日收盘价"]
            hold["当日"]=[trade_day for i in range(len(hold.index))]
            hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="当日")
            del hold["当日"]
            hold=copy.deepcopy(calcu_each_value())
            hold,blacklist,rest=check_blk()
            my_portfolio1[trade_day]["holdings"]=hold
            my_portfolio1[trade_day]["money"]=hold["价值"].sum()+rest
        if rest!=0:
            print("warning!!!!!!!!!!!!!!!!!!!!!!!!!!")
    return my_portfolio1
    
def adjust_according_to_MA2():
    def init_price_and_factor():
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
        return stockclose_sht,adjfac_sht
    def calcu_each_value():
        dex1=datelist[i-1]
        d1=datelist[i]
        hold1=copy.deepcopy(hold)
        for i1 in hold1.index:
            hold1.loc[i1,"价值"]
            hold_ex.loc[i1,"价值"]
            code1=hold1.loc[i1,"代码"]
            hold1.loc[i1,"价值"]=(
                                (stockclose_sht.loc[d1,code1]*adjfac_sht.loc[d1,code1]) /
                                (stockclose_sht.loc[dex1,code1]*adjfac_sht.loc[dex1,code1])
                                    )*hold_ex.loc[i1,"价值"]
        return hold1
    def check_blk():
        def calcu_MA(period=30):
            if datelist[datelist==trade_day].index[0]<=period:
                idx1=0
            else:
                idx1=datelist[datelist==trade_day].index[0]-period
            dtlis=datelist.iloc[idx1:idx1+period]
            MA1=(stockclose_sht.loc[dtlis,the_code]*adjfac_sht.loc[dtlis,the_code]).mean()
            return MA1
        def gojail():
            hold2=copy.deepcopy(hold1)
            rest_money2=hold2["价值"].iat[i1]
            hold2["价值"].iat[i1]=0
            t2=0
            for k in range(len(hold2.index)):
                if k!=i1 and not(hold2["代码"].iat[k] in blacklist1.keys() and (blacklist1[hold2["代码"].iat[k]]["status"]=="on_jail")):
                    t2=t2+1
            hold2["价值"].sum()
            return hold2,rest_money2+rest1
        def outjail():
            hold2=copy.deepcopy(hold1)
            t2=0
            for k in range(len(hold2.index)):
                if k!=i1 and hold2["价值"].iat[k]>0:
                    t2=t2+1
            rest_money2=(rest1)/(len(hold2.index)-t2)
            hold2["价值"].iat[i1]=rest_money2
            return hold2,rest1-rest_money2
        hold1=copy.deepcopy(hold)
        rest1=rest
        blacklist1=copy.deepcopy(blacklist)
        for i1 in range(len(hold.index)):
            the_code=hold["代码"].iat[i1]
            tmp1=stockclose_sht.loc[trade_day,the_code]
            tmp2=adjfac_sht.loc[trade_day,the_code]
            the_MA=calcu_MA(30)
            if the_code in blacklist1.keys():
                if tmp1*tmp2>=the_MA:
                    print("hero back!",i1,hold1["名称"].iat[i1],trade_day)
                    blacklist1.pop(the_code)
                    hold1,rest1=outjail()
            else:
                #不在黑名单
                if tmp1*tmp2<=the_MA:
                    print("it goes to jail!",i1,hold1["名称"].iat[i1],trade_day)
                    blacklist1[the_code]={}
                    blacklist1[the_code]["status"]="on_jail"
                    hold1,rest1=gojail()
        return hold1,blacklist1,rest1
    def do_adjust():
        h1=copy.deepcopy(hold)
        hex1=copy.deepcopy(hold_middle_ex)
        for i1 in h1.index:
            code1=h1.loc[i1,"代码"]
            if code1 in hex1["代码"].values:
                h1.loc[i1,"价值"]=hex1[hex1["代码"]==code1]["价值"].iat[0]
            else:
                h1.loc[i1,"价值"]=0.0
        r1=rest
        for i1 in hex1.index:
            code1=hex1.loc[i1,"代码"]
            if code1 in h1["代码"].values:
                pass
            else:
                r1=r1+hex1.loc[i1,"价值"]
        n1=0
        for i1 in h1.index:
            if h1.loc[i1,"价值"]==0.0:
                n1=n1+1
        rr1=r1
        for i1 in h1.index:
            code1=h1.loc[i1,"代码"]
            if code1 in hex1["代码"].values:
                pass
            else:
                h1.loc[i1,"价值"]=r1/n1
                rr1=rr1-h1.loc[i1,"价值"]
        hex1["价值"].sum()+rest
        h1["价值"].sum()+rr1
        return h1,rr1

        
        
        
        
    stockclose_sht,adjfac_sht=init_price_and_factor()
    adjust_date_list=pd.Series(
        list(my_portfolio.keys())).sort_values(ascending=True
                                               ).reset_index(drop=True)
    my_portfolio1=copy.deepcopy(my_portfolio)
    blacklist,rest={},0
    
    for i in range(len(datelist)):
    #for i in range(134): 
        #i=134,74,151
        trade_day=datelist[i]
        if trade_day<adjust_date_list[0]:
            continue
        print(i/len(datelist))
        if trade_day in adjust_date_list.values:
            if trade_day==adjust_date_list[0]:
                hold=copy.deepcopy(my_portfolio1[trade_day]["holdings"])
                hold["上一调仓日"]=[trade_day for i in range(len(hold.index))] 
                hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="上一调仓日")
                hold["当日复权因子"]=hold["调仓日复权因子"].copy()
                hold["当日收盘价"]=hold["调仓日收盘价"].copy()
                my_portfolio1[trade_day]["holdings"]=hold
                my_portfolio1[trade_day]["money"]=hold["价值"].sum()
            else:
                hold_ex=copy.deepcopy(my_portfolio1[datelist[i-1]]["holdings"])
                hold=copy.deepcopy(hold_ex)
                hold_middle_ex=copy.deepcopy(calcu_each_value())
                hold=copy.deepcopy(my_portfolio1[trade_day]["holdings"])
                hold["上一调仓日"]=[trade_day for i in range(len(hold.index))]
                hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="上一调仓日")
                hold["当日复权因子"]=hold["调仓日复权因子"].copy()
                hold["当日收盘价"]=hold["调仓日收盘价"].copy()
                hold["价值"].sum()
                hold,rest=do_adjust()
                blacklist=dict()
                my_portfolio1[trade_day]["holdings"]=hold
                my_portfolio1[trade_day]["money"]=hold["价值"].sum()+rest
        else:
            my_portfolio1[trade_day]=dict()
            hold_ex=copy.deepcopy(my_portfolio1[datelist[i-1]]["holdings"])
            hold=copy.deepcopy(hold_ex)
            del hold["当日复权因子"]
            del hold["当日收盘价"]
            hold["当日"]=[trade_day for i in range(len(hold.index))]
            hold=add_close_and_fac(hold,stockclose_dict,adjfactor_dict,on="当日")
            del hold["当日"]
            hold=copy.deepcopy(calcu_each_value())
            hold,blacklist,rest=check_blk()
            my_portfolio1[trade_day]["holdings"]=hold
            my_portfolio1[trade_day]["money"]=hold["价值"].sum()+rest
    return my_portfolio1
    
    

app=xw.App(visible=False,add_book=False)
app.display_alerts=False
app.screen_updating=False

dirpath=r"D:\desktop\mydatabase\stockclose"
stockclose_dict,filenamelis=read_one_after_one(dirpath)
stockclose_dict,eptlis=just_pile(stockclose_dict,ept=["null"])

dirpath=r"D:\desktop\mydatabase\adjfac"
adjfactor_dict,filenamelis=read_one_after_one(dirpath)
adjfactor_dict,eptlis=just_pile(adjfactor_dict,ept=["null"])

dirpath=r"D:\desktop\mydatabase\fundlove"
fundlove,filenamelis=read_one_after_one(dirpath,date=False)

del dirpath,filenamelis,eptlis

datelist=get_datelist([adjfactor_dict,stockclose_dict])
my_portfolio=dict()
for fundlove_filename in fundlove.keys():
    pass
    #fundlove_filename="20180331.xlsx"
    if fundlove_filename=="20221231.xlsx":
        break
    toplove=fundlove_get_top(fundlove[fundlove_filename],topnumber=50,use_hk=False)
    adjust_day=pd.to_datetime(fundlove_filename.split(".")[0])
    adjust_day=search_adjust_day(adjust_day)
    toplove["价值"]=[1/50 for i in range(len(toplove.index))]
    my_portfolio[adjust_day]={"holdings":toplove}
del toplove,fundlove_filename



change_mode=3
if change_mode==0:
    my_portfolio=adjust_according_to_price()
elif change_mode==1:
    my_portfolio=adjust_according_to_loverank()
elif change_mode==2:
    my_portfolio=adjust_according_to_MA()
elif change_mode==3:
    my_portfolio=adjust_according_to_MA2()  #每份单独使用，空仓就空仓、不补给其他票

#集成总表和详表
datelist=pd.Series(list(my_portfolio.keys())).sort_values(ascending=True).reset_index(drop=True)
general_df=pd.DataFrame(columns=["日期","每日净值"])
detailed_df=pd.DataFrame()
for i in range(len(datelist)):
    trade_day=datelist[i]
    money=my_portfolio[trade_day]["money"]
    tmp1=my_portfolio[trade_day]["holdings"].copy()
    tmp1.index=[trade_day for j in range(len(tmp1.index))]
    detailed_df=pd.concat([detailed_df,tmp1],axis=0)
    tmp2=pd.DataFrame((trade_day,money),index=["日期","每日净值"],columns=[i]).T
    general_df=pd.concat([general_df,tmp2],axis=0)
    
#计算最大回撤
general_df["accu_max"]=[0.0 for i in range(len(general_df.index))]
for i in range(len(general_df.index)):
    tmp1=general_df.iloc[:i,1].max()
    if (tmp1==tmp1):
        general_df["accu_max"].iat[i]=general_df["每日净值"].iat[i]/tmp1-1
general_df["accu_max"].min()

#输出表格
general_df.to_excel(r'D:\desktop\fundlove\总表.xlsx',header=True,index=True)
detailed_df.to_excel(r'D:\desktop\fundlove\详表.xlsx',header=True,index=True)        





app.kill()











