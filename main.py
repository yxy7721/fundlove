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
import sys
sys.path.append(r"D:\desktop\fundlove") 
import fundlove2 as y2
#import fundlove as y1
import fundlove3 as y3

if __name__ == "__main__":
    
    app=xw.App(visible=False,add_book=False)
    app.display_alerts=False
    app.screen_updating=False
    
    #第一步准备数据并清洗
    prepare=y2.PrepareData([r"D:\desktop\mydatabase\stockclose",
                            r"D:\desktop\mydatabase\adjfac",
                            r"D:\desktop\mydatabase\fundlove"],
                           app
                           )
    dataset=prepare.make_data()
    del prepare
    
    #第二步筛选每日持仓（暂时不考虑每日净值）
    #MyPortfolio、MyPortfolioDueFundsNum、MyPortfolioDueAssetRatio、MyPortfolioDueFundsChange
    #MyPortfolioNoMA、MyPortfolioDueFundsNumNoMA、MyPortfolioDueAssetRatioNoMA、MyPortfolioDueFundsChangeNoMA 
    my=y2.MyPortfolioDueAssetRatioNoMA(dataset,20) 
    #my.init_price_and_factor()
    basic_parameter=my.get_para_data()
    my_portfolio=my.get_pf()
    del my
    
    #第三步计算每日净值
    cn=y2.CalcuNav(my_portfolio,basic_parameter)
    my_portfolio=cn.do_and_return()
    
    #另一种思路的第二步，生成双因子
    
    
    
    
    
    
    
    



    #集成总表和详表
    datelist=pd.Series(list(my_portfolio.keys())).sort_values(ascending=True).reset_index(drop=True)
    general_df=pd.DataFrame(columns=["日期","每日净值"])
    detailed_df=pd.DataFrame()
    for i in range(len(datelist)):
        trade_day=datelist[i]
        money=my_portfolio[trade_day]["价值"].sum()
        tmp1=my_portfolio[trade_day].copy()
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











