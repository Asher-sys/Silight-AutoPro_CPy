# -*- coding: utf-8 -*-

"""
Module implementing PIForm.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget

from Ui_PIForm import Ui_Form
import pandas as pd
import os

class PIForm(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    df1 = []#全局变量定义
    df3_new = []
    df3 = []
    df3_XY = []
    df3_cnt = []
    
    df3_ICHx_mean_value = []
    
    
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget (defaults to None)
        @type QWidget (optional)
        """
        super(PIForm, self).__init__(parent)
        self.setupUi(self)
        
    '''获取CSV格式原始数据'''
    def Get_csv(self):
        global df1
        inputfile = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\mydlg.csv'
        outputfile = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\mydlg_out.csv'
        df1 = pd.read_csv(inputfile,encoding='utf-8',header=None,sep=None,engine='python')
#        print('处理对象:%s'%inputfile)
        df1.to_csv(outputfile)
        
        '''保留有效数据部分，并重命名列'''
    def Get_New_csv(self):
        global df1, df3,  df3_new
        df_cols = df1.iloc[14]
        df2 = df1.rename(columns = df_cols)
        Leak_cols_num = [i for i,x in enumerate(df2.columns) if x.find('Leak')!=-1]
        df3 = df2.drop(df2.columns[Leak_cols_num],axis = 1)
        df3_new = df3.drop([0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17],axis = 0)
        df3_new = df3_new.drop(' ',axis = 1)
        df3_new = df3_new.rename(columns={'Parm_Name':'Bin'})#更改列名
        
        '''判断一张晶圆上有多少颗芯片'''
    def Wafer_Dies(self):
        global df3_new, df3_XY, df3_cnt
        df3_cnt = df3_new[df3_new.columns[0]].count()
        df3_XY = df3_new
        if df3_cnt == 46:
            XCols = ['-4',\
             '-3','-3','-3','-3','-3',\
             '-2','-2','-2','-2','-2','-2','-2',\
             '-1','-1','-1','-1','-1','-1','-1',\
             '0','0','0','0','0','0','0',\
             '1','1','1','1','1','1','1',\
             '2','2','2','2','2','2','2',\
             '3','3','3','3','3']
            YCols = ['0',\
             '2','1','0','-1','-2',\
             '-3','-2','-1','0','1','2','3',\
             '3','2','1','0','-1','-2','-3',\
             '-3','-2','-1','0','1','2','3',\
             '3','2','1','0','-1','-2','-3',\
             '-3','-2','-1','0','1','2','3',\
             '2','1','0','-1','-2']
            df3_XY.insert(1,'X',XCols) #插入横坐标X
            df3_XY.insert(2,'Y',YCols) #插入纵坐标Y
            
        elif df3_cnt == 100:
            XCols = ['-3','-3','-3','-3','-3','-3','-3','-3','-3','-3',\
             '-2','-2','-2','-2','-2','-2','-2','-2','-2','-2','-2','-2',\
             '-1','-1','-1','-1','-1','-1','-1','-1','-1','-1','-1','-1','-1','-1','-1','-1',\
             '0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0',\
             '1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1',\
             '2','2','2','2','2','2','2','2','2','2','2','2','2','2',\
             '3','3','3','3','3','3','3','3','3','3','3','3',\
             '4','4','4','4']
            YCols = ['5','4','3','2','1','0','-1','-2','-3','-4',\
             '-5','-4','-3','-2','-1','0','1','2','3','4','5','6',\
             '8','7','6','5','4','3','2','1','0','-1','-2','-3','-4','-5','-6','-7',\
             '-7','-6','-5','-4','-3','-2','-1','0','1','2','3','4','5','6','7','8',\
             '8','7','6','5','4','3','2','1','0','-1','-2','-3','-4','-5','-6','-7',\
             '-6','-5','-4','-3','-2','-1','0','1','2','3','4','5','6','7',\
             '6','5','4','3','2','1','0','-1','-2','-3','-4','-5',\
             '-1','0','1','2']
            df3_XY.insert(1,'X',XCols)
            df3_XY.insert(2,'Y',YCols)
        else:
            pass
            
    '''获取芯片各通道正极电流原始数据'''
    def Get_PI_data(self):
        global df3_new, df3_ICHx_mean_value
        PI_threh = 20.000 #正极电流有效值阈值
        ICH_cols = ['ICH1','ICH2','ICH3','ICH4','ICH5','ICH6','ICH7','ICH8',\
            'ICH9','ICH10','ICH11','ICH12','ICH13','ICH14','ICH15','ICH16',\
            'ICH17','ICH18','ICH19','ICH20','ICH21','ICH22','ICH23','ICH24',\
            'ICH25','ICH26','ICH27','ICH28','ICH29','ICH30','ICH31','ICH32',\
            'ICH33','ICH34','ICH35','ICH36','ICH37','ICH38','ICH39','ICH40',\
            'ICH41','ICH42','ICH43','ICH44','ICH45','ICH46','ICH47','ICH48',\
            'ICH49','ICH50','ICH51','ICH52']
        
        #删除正极电流小于20mA的列
        for ICHcols_num in ICH_cols:
            df3_new[ICHcols_num] = df3_new[ICHcols_num].astype('float64')#类型转换
            df3_ICH = df3_new[df3_new[ICHcols_num] > PI_threh]
            
        PI_PATH = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\PI20_out.csv'
        df3_ICH.to_csv(PI_PATH)
            
        #同颗芯片52通道正极电流汇总统计,正极电流数据大于0.0有效
        df3_ICHx_new = df3_new[ICH_cols]
        df3_ICHx = df3_ICHx_new.loc[:,(df3_ICHx_new > 0.0).all(axis=0)]
        ICHx_index = df3_ICHx_new.index-18
        df3_ICHx_Index = []
        df3_ICHx_mean = []
        for ICHxindex_num in ICHx_index:
            df3_ICHx_Index = df3_ICHx.iloc[ICHxindex_num]
            df3_ICHx_mean.append(df3_ICHx_Index.mean())
            df3_ICHx_mean = [round(i,1) for i in df3_ICHx_mean]#list取1位小数
            
        df3_ICHx_mean_value = df3_ICHx_mean
            
    
    @pyqtSlot()
    def on_btn_PI_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
#        print("点击正极电流测试项按钮")
        self.Get_csv()
        self.Get_New_csv()
        self.Wafer_Dies()
        self.Get_PI_data()
        print("正极电流测试项完成!")
