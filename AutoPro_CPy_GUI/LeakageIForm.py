# -*- coding: utf-8 -*-

"""
Module implementing LeakageIForm.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget

from Ui_LeakageIForm import Ui_Form
import pandas as pd
import os
#import xlwt

class LeakageIForm(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    df1 = []#全局变量定义
    df3_new = []
    df3 = []
    df3_XY = []
    df3_cnt = []
    df_Leak = []
    
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget (defaults to None)
        @type QWidget (optional)
        """
        super(LeakageIForm, self).__init__(parent)
        self.setupUi(self)
        
    '''获取CSV格式原始数据'''
    def Get_csv(self):
        global df1
        inputfile = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\mydlg.csv'
        outputfile = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\mydlg_out.csv'
#        inputfile = os.path.dirname(os.path.realpath('__file__'))+'\csv_data\\mydlg.csv'
#        outputfile = os.path.dirname(os.path.realpath('__file__'))+'\csv_data\\mydlg_out.csv'
        df1 = pd.read_csv(inputfile,encoding='utf-8',header=None,sep=None,engine='python')
#        print('处理对象:%s'%inputfile)
        df1.to_csv(outputfile)
        
        '''保留有效数据部分，并重命名列'''
    def Get_New_csv(self):
        global df1, df3,  df3_new, df_Leak
        df_cols = df1.iloc[14]
        df2 = df1.rename(columns = df_cols)
        Leak_cols_num = [i for i,x in enumerate(df2.columns) if x.find('Leak')!=-1]
        df3 = df2.drop(df2.columns[Leak_cols_num],axis = 1)
        df3_new = df3.drop([0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17],axis = 0)
        df3_new = df3_new.drop(' ',axis = 1)
        df3_new = df3_new.rename(columns={'Parm_Name':'Bin'})#更改列名
        
        #获取漏电流数据.csv格式
        Leak_cols = df2.columns[Leak_cols_num]
        df_Leak = df2[Leak_cols]
        
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
            
    '''获取漏电流原始数据'''
    def Get_LeakI_data(self):
        global df_Leak
        LeakI_PATH = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\LeakI_out.csv'
#        LeakI_PATH = os.path.dirname(os.path.realpath('__file__'))+'\csv_data\\LeakI_out.csv'
        df_Leak.to_csv(LeakI_PATH)
        
    
    @pyqtSlot()
    def on_btn_LeakageI_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
#        print("点击漏电流测试项按钮")
        self.Get_csv()
        self.Get_New_csv()
        self.Wafer_Dies()
        self.Get_LeakI_data()
        print("漏电流测试项完成!")
