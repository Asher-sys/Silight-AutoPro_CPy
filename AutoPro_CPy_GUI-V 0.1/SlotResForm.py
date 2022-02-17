# -*- coding: utf-8 -*-

"""
Module implementing SlotResForm.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget

from Ui_SlotResForm import Ui_Form
import pandas as pd
import os
import xlwt

class SlotResForm(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    df1 = []#全局变量定义
    df3_new = []
    df3 = []
    df3_XY = []
    df3_cnt = []
    df3_RCH = []
    df3_mRCHr_new = []
    
    df3_RCH1_mean_value = []
    df3_RCH2_mean_value = []
    df3_RCH3_mean_value = []
    df3_RCH4_mean_value = []
    df3_RCH5_mean_value = []
    
    df3_RCH1_XY = []
    df3_RCH2_XY = []
    df3_RCH3_XY = []
    df3_RCH4_XY = []
    df3_RCH5_XY = []
    
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget (defaults to None)
        @type QWidget (optional)
        """
        super(SlotResForm, self).__init__(parent)
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
            
    '''获取平板区测试数据'''
    def Get_RCHSlot(self):
        global df3,  df3_new, df3_RCH
        RCH_cols_num = [i for i,x in enumerate(df3.columns) if x.find('RCH')!=-1]
        RCH_cols = df3.columns[RCH_cols_num]
        for RCHcols_num in RCH_cols:
            df3_new[RCHcols_num] = df3_new[RCHcols_num].astype('float64')#类型转换
        
        df3_RCH = df3_new[RCH_cols]
        
        PATH_Save = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCHx_out.csv'
        df3_RCH.to_csv(PATH_Save)
    
    '''从原始数据中提取RCH1、RCH2、RCH3、RCH4、RCH5项,并计算出电阻值,得到电阻值表格'''
    def Get_RCHx_Data_csv(self):
        global df3_RCH,  df3_mRCHr_new
        mRCH1_rcols_num = [i for i,x in enumerate(df3_RCH.columns) if x.find('RCH1')!=-1]
        mRCH2_rcols_num = [i for i,x in enumerate(df3_RCH.columns) if x.find('RCH2')!=-1]
        mRCH3_rcols_num = [i for i,x in enumerate(df3_RCH.columns) if x.find('RCH3')!=-1]
        mRCH4_rcols_num = [i for i,x in enumerate(df3_RCH.columns) if x.find('RCH4')!=-1]
        mRCH5_rcols_num = [i for i,x in enumerate(df3_RCH.columns) if x.find('RCH5')!=-1]
        mRCH1_rcols = df3_RCH.columns[mRCH1_rcols_num]
        mRCH2_rcols = df3_RCH.columns[mRCH2_rcols_num]
        mRCH3_rcols = df3_RCH.columns[mRCH3_rcols_num]
        mRCH4_rcols = df3_RCH.columns[mRCH4_rcols_num]
        mRCH5_rcols = df3_RCH.columns[mRCH5_rcols_num]
        df3_mRCH1 = df3_RCH[mRCH1_rcols]
        df3_mRCH2 = df3_RCH[mRCH2_rcols]
        df3_mRCH3 = df3_RCH[mRCH3_rcols]
        df3_mRCH4 = df3_RCH[mRCH4_rcols]
        df3_mRCH5 = df3_RCH[mRCH5_rcols]
        df3_mRxCH_1 = []; df3_mRxCH_2 = []; df3_mRxCH_3 = []; df3_mRxCH_4 = []; df3_mRxCH_5 = []
        df3_mRxCH_6 = []; df3_mRxCH_7 = []; df3_mRxCH_8 = []; df3_mRxCH_9 = []; df3_mRxCH_10 = []
        
        df3_mRxCH1 = []; df3_mRxCH2 = []; df3_mRxCH3 = []; df3_mRxCH4 = []; df3_mRxCH5 = []
        df3_mRxCH_1 = (df3_mRCH1[df3_mRCH1.columns[1]] - df3_mRCH1[df3_mRCH1.columns[0]]) / 0.005
        df3_mRxCH_2 = (df3_mRCH1[df3_mRCH1.columns[2]] - df3_mRCH1[df3_mRCH1.columns[1]]) / 0.005
        df3_mRxCH_3 = (df3_mRCH1[df3_mRCH1.columns[3]] - df3_mRCH1[df3_mRCH1.columns[2]]) / 0.005
        df3_mRxCH_4 = (df3_mRCH1[df3_mRCH1.columns[4]] - df3_mRCH1[df3_mRCH1.columns[3]]) / 0.005
        df3_mRxCH_5 = (df3_mRCH1[df3_mRCH1.columns[5]] - df3_mRCH1[df3_mRCH1.columns[4]]) / 0.005
        df3_mRxCH_6 = (df3_mRCH1[df3_mRCH1.columns[6]] - df3_mRCH1[df3_mRCH1.columns[5]]) / 0.005
        df3_mRxCH_7 = (df3_mRCH1[df3_mRCH1.columns[7]] - df3_mRCH1[df3_mRCH1.columns[6]]) / 0.005
        df3_mRxCH_8 = (df3_mRCH1[df3_mRCH1.columns[8]] - df3_mRCH1[df3_mRCH1.columns[7]]) / 0.005
        df3_mRxCH_9 = (df3_mRCH1[df3_mRCH1.columns[9]] - df3_mRCH1[df3_mRCH1.columns[8]]) / 0.005
        df3_mRxCH_10 = (df3_mRCH1[df3_mRCH1.columns[10]] - df3_mRCH1[df3_mRCH1.columns[9]]) / 0.005
        df3_mRxCH1 = pd.DataFrame({'RCH1_1':df3_mRxCH_1,'RCH1_2':df3_mRxCH_2, 'RCH1_3':df3_mRxCH_3,\
                           'RCH1_4':df3_mRxCH_4,'RCH1_5':df3_mRxCH_5, 'RCH1_6':df3_mRxCH_6,\
                           'RCH1_7':df3_mRxCH_7,'RCH1_8':df3_mRxCH_8, 'RCH1_9':df3_mRxCH_9,\
                           'RCH1_10':df3_mRxCH_10})
                           
        df3_mRxCH_1 = (df3_mRCH2[df3_mRCH2.columns[1]] - df3_mRCH2[df3_mRCH2.columns[0]]) / 0.005
        df3_mRxCH_2 = (df3_mRCH2[df3_mRCH2.columns[2]] - df3_mRCH2[df3_mRCH2.columns[1]]) / 0.005
        df3_mRxCH_3 = (df3_mRCH2[df3_mRCH2.columns[3]] - df3_mRCH2[df3_mRCH2.columns[2]]) / 0.005
        df3_mRxCH_4 = (df3_mRCH2[df3_mRCH2.columns[4]] - df3_mRCH2[df3_mRCH2.columns[3]]) / 0.005
        df3_mRxCH_5 = (df3_mRCH2[df3_mRCH2.columns[5]] - df3_mRCH2[df3_mRCH2.columns[4]]) / 0.005
        df3_mRxCH_6 = (df3_mRCH2[df3_mRCH2.columns[6]] - df3_mRCH2[df3_mRCH2.columns[5]]) / 0.005
        df3_mRxCH_7 = (df3_mRCH2[df3_mRCH2.columns[7]] - df3_mRCH2[df3_mRCH2.columns[6]]) / 0.005
        df3_mRxCH_8 = (df3_mRCH2[df3_mRCH2.columns[8]] - df3_mRCH2[df3_mRCH2.columns[7]]) / 0.005
        df3_mRxCH_9 = (df3_mRCH2[df3_mRCH2.columns[9]] - df3_mRCH2[df3_mRCH2.columns[8]]) / 0.005
        df3_mRxCH_10 = (df3_mRCH2[df3_mRCH2.columns[10]] - df3_mRCH2[df3_mRCH2.columns[9]]) / 0.005
        df3_mRxCH2 = pd.DataFrame({'RCH2_1':df3_mRxCH_1,'RCH2_2':df3_mRxCH_2, 'RCH2_3':df3_mRxCH_3,\
                           'RCH2_4':df3_mRxCH_4,'RCH2_5':df3_mRxCH_5, 'RCH2_6':df3_mRxCH_6,\
                           'RCH2_7':df3_mRxCH_7,'RCH2_8':df3_mRxCH_8, 'RCH2_9':df3_mRxCH_9,\
                           'RCH2_10':df3_mRxCH_10})
                           
        df3_mRxCH_1 = (df3_mRCH3[df3_mRCH3.columns[1]] - df3_mRCH3[df3_mRCH3.columns[0]]) / 0.005
        df3_mRxCH_2 = (df3_mRCH3[df3_mRCH3.columns[2]] - df3_mRCH3[df3_mRCH3.columns[1]]) / 0.005
        df3_mRxCH_3 = (df3_mRCH3[df3_mRCH3.columns[3]] - df3_mRCH3[df3_mRCH3.columns[2]]) / 0.005
        df3_mRxCH_4 = (df3_mRCH3[df3_mRCH3.columns[4]] - df3_mRCH3[df3_mRCH3.columns[3]]) / 0.005
        df3_mRxCH_5 = (df3_mRCH3[df3_mRCH3.columns[5]] - df3_mRCH3[df3_mRCH3.columns[4]]) / 0.005
        df3_mRxCH_6 = (df3_mRCH3[df3_mRCH3.columns[6]] - df3_mRCH3[df3_mRCH3.columns[5]]) / 0.005
        df3_mRxCH_7 = (df3_mRCH3[df3_mRCH3.columns[7]] - df3_mRCH3[df3_mRCH3.columns[6]]) / 0.005
        df3_mRxCH_8 = (df3_mRCH3[df3_mRCH3.columns[8]] - df3_mRCH3[df3_mRCH3.columns[7]]) / 0.005
        df3_mRxCH_9 = (df3_mRCH3[df3_mRCH3.columns[9]] - df3_mRCH3[df3_mRCH3.columns[8]]) / 0.005
        df3_mRxCH_10 = (df3_mRCH3[df3_mRCH3.columns[10]] - df3_mRCH3[df3_mRCH3.columns[9]]) / 0.005
        df3_mRxCH3 = pd.DataFrame({'RCH3_1':df3_mRxCH_1,'RCH3_2':df3_mRxCH_2, 'RCH3_3':df3_mRxCH_3,\
                           'RCH3_4':df3_mRxCH_4,'RCH3_5':df3_mRxCH_5, 'RCH3_6':df3_mRxCH_6,\
                           'RCH3_7':df3_mRxCH_7,'RCH3_8':df3_mRxCH_8, 'RCH3_9':df3_mRxCH_9,\
                           'RCH3_10':df3_mRxCH_10})
                           
        df3_mRxCH_1 = (df3_mRCH4[df3_mRCH4.columns[1]] - df3_mRCH4[df3_mRCH4.columns[0]]) / 0.005
        df3_mRxCH_2 = (df3_mRCH4[df3_mRCH4.columns[2]] - df3_mRCH4[df3_mRCH4.columns[1]]) / 0.005
        df3_mRxCH_3 = (df3_mRCH4[df3_mRCH4.columns[3]] - df3_mRCH4[df3_mRCH4.columns[2]]) / 0.005
        df3_mRxCH_4 = (df3_mRCH4[df3_mRCH4.columns[4]] - df3_mRCH4[df3_mRCH4.columns[3]]) / 0.005
        df3_mRxCH_5 = (df3_mRCH4[df3_mRCH4.columns[5]] - df3_mRCH4[df3_mRCH4.columns[4]]) / 0.005
        df3_mRxCH_6 = (df3_mRCH4[df3_mRCH4.columns[6]] - df3_mRCH4[df3_mRCH4.columns[5]]) / 0.005
        df3_mRxCH_7 = (df3_mRCH4[df3_mRCH4.columns[7]] - df3_mRCH4[df3_mRCH4.columns[6]]) / 0.005
        df3_mRxCH_8 = (df3_mRCH4[df3_mRCH4.columns[8]] - df3_mRCH4[df3_mRCH4.columns[7]]) / 0.005
        df3_mRxCH_9 = (df3_mRCH4[df3_mRCH4.columns[9]] - df3_mRCH4[df3_mRCH4.columns[8]]) / 0.005
        df3_mRxCH_10 = (df3_mRCH4[df3_mRCH4.columns[10]] - df3_mRCH4[df3_mRCH4.columns[9]]) / 0.005
        df3_mRxCH4 = pd.DataFrame({'RCH4_1':df3_mRxCH_1,'RCH4_2':df3_mRxCH_2, 'RCH4_3':df3_mRxCH_3,\
                           'RCH4_4':df3_mRxCH_4,'RCH4_5':df3_mRxCH_5, 'RCH4_6':df3_mRxCH_6,\
                           'RCH4_7':df3_mRxCH_7,'RCH4_8':df3_mRxCH_8, 'RCH4_9':df3_mRxCH_9,\
                           'RCH4_10':df3_mRxCH_10})
                           
        df3_mRxCH_1 = (df3_mRCH5[df3_mRCH5.columns[1]] - df3_mRCH5[df3_mRCH5.columns[0]]) / 0.005
        df3_mRxCH_2 = (df3_mRCH5[df3_mRCH5.columns[2]] - df3_mRCH5[df3_mRCH5.columns[1]]) / 0.005
        df3_mRxCH_3 = (df3_mRCH5[df3_mRCH5.columns[3]] - df3_mRCH5[df3_mRCH5.columns[2]]) / 0.005
        df3_mRxCH_4 = (df3_mRCH5[df3_mRCH5.columns[4]] - df3_mRCH5[df3_mRCH5.columns[3]]) / 0.005
        df3_mRxCH_5 = (df3_mRCH5[df3_mRCH5.columns[5]] - df3_mRCH5[df3_mRCH5.columns[4]]) / 0.005
        df3_mRxCH_6 = (df3_mRCH5[df3_mRCH5.columns[6]] - df3_mRCH5[df3_mRCH5.columns[5]]) / 0.005
        df3_mRxCH_7 = (df3_mRCH5[df3_mRCH5.columns[7]] - df3_mRCH5[df3_mRCH5.columns[6]]) / 0.005
        df3_mRxCH_8 = (df3_mRCH5[df3_mRCH5.columns[8]] - df3_mRCH5[df3_mRCH5.columns[7]]) / 0.005
        df3_mRxCH_9 = (df3_mRCH5[df3_mRCH5.columns[9]] - df3_mRCH5[df3_mRCH5.columns[8]]) / 0.005
        df3_mRxCH_10 = (df3_mRCH5[df3_mRCH5.columns[10]] - df3_mRCH5[df3_mRCH5.columns[9]]) / 0.005
        df3_mRxCH5 = pd.DataFrame({'RCH5_1':df3_mRxCH_1,'RCH5_2':df3_mRxCH_2, 'RCH5_3':df3_mRxCH_3,\
                           'RCH5_4':df3_mRxCH_4,'RCH5_5':df3_mRxCH_5, 'RCH5_6':df3_mRxCH_6,\
                           'RCH5_7':df3_mRxCH_7,'RCH5_8':df3_mRxCH_8, 'RCH5_9':df3_mRxCH_9,\
                           'RCH5_10':df3_mRxCH_10})
                           
        df3_mRCHr_new = pd.concat([df3_mRxCH1, df3_mRxCH2, df3_mRxCH3, df3_mRxCH4, df3_mRxCH5], axis=1)
        PATH_Save1 = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\mRCHx_out.csv'
        df3_mRCHr_new.to_csv(PATH_Save1)
        
        '''获取同颗芯片不同通道电阻平均值'''
    def Get_RCH_Mean(self):
        global df3_mRCHr_new
        global df3_RCH1_mean_value, df3_RCH2_mean_value, df3_RCH3_mean_value
        global df3_RCH4_mean_value, df3_RCH5_mean_value
        df3_RCHr_new = df3_mRCHr_new
        RCH1_rcols_num = [i for i,x in enumerate(df3_RCHr_new.columns) if x.find('RCH1')!=-1]
        RCH2_rcols_num = [i for i,x in enumerate(df3_RCHr_new.columns) if x.find('RCH2')!=-1]
        RCH3_rcols_num = [i for i,x in enumerate(df3_RCHr_new.columns) if x.find('RCH3')!=-1]
        RCH4_rcols_num = [i for i,x in enumerate(df3_RCHr_new.columns) if x.find('RCH4')!=-1]
        RCH5_rcols_num = [i for i,x in enumerate(df3_RCHr_new.columns) if x.find('RCH5')!=-1]
        
        RCH1_rcols = df3_RCHr_new.columns[RCH1_rcols_num]
        RCH2_rcols = df3_RCHr_new.columns[RCH2_rcols_num]
        RCH3_rcols = df3_RCHr_new.columns[RCH3_rcols_num]
        RCH4_rcols = df3_RCHr_new.columns[RCH4_rcols_num]
        RCH5_rcols = df3_RCHr_new.columns[RCH5_rcols_num]
        
        df3_RCH1 = df3_RCHr_new[RCH1_rcols]
        df3_RCH2 = df3_RCHr_new[RCH2_rcols]
        df3_RCH3 = df3_RCHr_new[RCH3_rcols]
        df3_RCH4 = df3_RCHr_new[RCH4_rcols]
        df3_RCH5 = df3_RCHr_new[RCH5_rcols]
        
        df3_RCH1_index = [];df3_RCH2_index = [];df3_RCH3_index = [];df3_RCH4_index = [];df3_RCH5_index = []
        df3_RCH1_mean = [];df3_RCH2_mean = [];df3_RCH3_mean = [];df3_RCH4_mean = [];df3_RCH5_mean = []
        
        RCHx_index = df3_RCH1.index  -18
        for RCHx_index_num in RCHx_index:
            df3_RCH1_index = df3_RCH1.iloc[RCHx_index_num]
            df3_RCH2_index = df3_RCH2.iloc[RCHx_index_num]
            df3_RCH3_index = df3_RCH3.iloc[RCHx_index_num]
            df3_RCH4_index = df3_RCH4.iloc[RCHx_index_num]
            df3_RCH5_index = df3_RCH5.iloc[RCHx_index_num]
            df3_RCH1_mean.append(df3_RCH1_index.mean())
            df3_RCH2_mean.append(df3_RCH2_index.mean())
            df3_RCH3_mean.append(df3_RCH3_index.mean())
            df3_RCH4_mean.append(df3_RCH4_index.mean())
            df3_RCH5_mean.append(df3_RCH5_index.mean())
            df3_RCH1_mean = [round(i,1) for i in df3_RCH1_mean]
            df3_RCH2_mean = [round(i,1) for i in df3_RCH2_mean]
            df3_RCH3_mean = [round(i,1) for i in df3_RCH3_mean]
            df3_RCH4_mean = [round(i,1) for i in df3_RCH4_mean]
            df3_RCH5_mean = [round(i,1) for i in df3_RCH5_mean]
        
        df3_RCH1_mean_value = df3_RCH1_mean
        df3_RCH2_mean_value = df3_RCH2_mean
        df3_RCH3_mean_value = df3_RCH3_mean
        df3_RCH4_mean_value = df3_RCH4_mean
        df3_RCH5_mean_value = df3_RCH5_mean
            
    '''获取含坐标的电阻平均值表格'''
    def Get_XYRCH1_Mean_csv(self):
        global df3_XY, df3_RCH1_XY, df3_RCH1_mean_value
        df3_XYRCH1 = df3_XY
        RCH1_mean = df3_RCH1_mean_value
        df3_XYRCH1.insert(3,'RCH1_mean',RCH1_mean)
        df3_RCH1_XY = df3_XYRCH1[df3_XYRCH1.columns[0:4]] #获取前4列['Bin','X','Y','RCH_mean'],根据需求修改
        
    def Get_XYRCH2_Mean_csv(self):
        global df3_XY,  df3_RCH2_XY, df3_RCH2_mean_value
        df3_XYRCH2 = df3_XY
        RCH2_mean = df3_RCH2_mean_value
        df3_XYRCH2.insert(3,'RCH2_mean',RCH2_mean)
        df3_RCH2_XY = df3_XYRCH2[df3_XYRCH2.columns[0:4]]
        
    def Get_XYRCH3_Mean_csv(self):
        global df3_XY,  df3_RCH3_XY, df3_RCH3_mean_value
        df3_XYRCH3 = df3_XY
        RCH3_mean = df3_RCH3_mean_value
        df3_XYRCH3.insert(3,'RCH3_mean',RCH3_mean)
        df3_RCH3_XY = df3_XYRCH3[df3_XYRCH3.columns[0:4]]
        
    def Get_XYRCH4_Mean_csv(self):
        global df3_XY,  df3_RCH4_XY, df3_RCH4_mean_value
        df3_XYRCH4 = df3_XY
        RCH4_mean = df3_RCH4_mean_value
        df3_XYRCH4.insert(3,'RCH4_mean',RCH4_mean)
        df3_RCH4_XY = df3_XYRCH4[df3_XYRCH4.columns[0:4]]
        
    def Get_XYRCH5_Mean_csv(self):
        global df3_XY,  df3_RCH5_XY, df3_RCH5_mean_value
        df3_XYRCH5 = df3_XY
        RCH5_mean = df3_RCH5_mean_value
        df3_XYRCH5.insert(3,'RCH5_mean',RCH5_mean)
        df3_RCH5_XY = df3_XYRCH5[df3_XYRCH5.columns[0:4]]
        
    '''得到非负坐标Map图'''
    def Get_XYRCH1New_Map(self):
        global df3_cnt, df3_RCH1_XY
        
        #颜色定义
        styleLightGreenBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')
        styleYellowBkg = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        styleLightBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;')
        styleIceBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue;')
        styleSkyBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;')
        styleOrangeBkg = xlwt.easyxf('pattern: pattern solid, fore_colour orange;')
        stylePaleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;')
        
        if df3_cnt == 46:
            arr1 = list([4 for x in range(46)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH1_XY['X'].tolist()
            RCH_Y = df3_RCH1_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(46):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(46):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        elif df3_cnt == 100:
            arr1 = list([8 for x in range(100)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH1_XY['X'].tolist()
            RCH_Y = df3_RCH1_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(100):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(100):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        else:
            pass
            
        df3_RCH1_XY.insert(3,'RCH_X',RCH_Xarrlist)
        df3_RCH1_XY.insert(4,'RCH_Y',RCH_Yarrlist)
        
        df3_RCH1_XY = df3_RCH1_XY[df3_RCH1_XY.columns[3:6]]
        
        #删除上次生成的文件,避免出现无法修改数据错误
        RCH_PATH = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH1-Mapdata.txt'
        RCH_PATH1 = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH1-Mapdata.xls'
        if os.path.exists(RCH_PATH):
            os.remove(RCH_PATH)
        elif os.path.exists(RCH_PATH1):
            os.remove(RCH_PATH1)
        else:
            pass
            
        #存储为txt文件
        with open(RCH_PATH,'a+',encoding='utf-8') as f1:
            for line in df3_RCH1_XY.values:
                f1.write((str(line[0])+'\t'+str(line[1])+'\t'+str(line[2])+'\n'))
        workbook1 = xlwt.Workbook()
        worksheet1 = workbook1.add_sheet('RCH1-Map')
        
        #txt文件地址
        filein  = RCH_PATH
        readfile  = open(filein,'r')
        for line in readfile:
            word = line.split()
            if len(word) == 3:
                y=int(word[1])
                x=int(word[0])
                RCH = word[2]
                if 1:
                    if (x == 5):    
                        worksheet1.write(x+7,y,RCH,stylePaleBlueBkg)
                    elif (x == 6):
                        worksheet1.write(x+5,y,RCH,styleLightGreenBkg)
                    elif (x == 7):
                        worksheet1.write(x+3,y,RCH,styleSkyBlueBkg)
                    elif (x == 8):
                        worksheet1.write(x+1,y,RCH,styleYellowBkg)
                    elif (x == 9):
                        worksheet1.write(x-1,y,RCH,styleLightBlueBkg)
                    elif (x == 10):
                        worksheet1.write(x-3,y,RCH,styleIceBlueBkg)
                    elif (x == 11):
                        worksheet1.write(x-5,y,RCH,styleOrangeBkg)
                    elif (x == 12):
                        worksheet1.write(x-7,y,RCH,stylePaleBlueBkg)
                    else:
                        worksheet1.write(x,y,RCH,styleOrangeBkg)
        for ix in range(256):
            worksheet1.col(ix).width =  256 * (5 + 1)
            
        workbook1.save(RCH_PATH1)
        print("晶圆RCH1电阻分布MAP完成!")
        
        
        
    def Get_XYRCH2New_Map(self):
        global df3_cnt, df3_RCH2_XY
        
        #颜色定义
        styleLightGreenBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')
        styleYellowBkg = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        styleLightBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;')
        styleIceBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue;')
        styleSkyBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;')
        styleOrangeBkg = xlwt.easyxf('pattern: pattern solid, fore_colour orange;')
        stylePaleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;')
        
        if df3_cnt == 46:
            arr1 = list([4 for x in range(46)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH2_XY['X'].tolist()
            RCH_Y = df3_RCH2_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(46):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(46):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        elif df3_cnt == 100:
            arr1 = list([8 for x in range(100)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH2_XY['X'].tolist()
            RCH_Y = df3_RCH2_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(100):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(100):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        else:
            pass
            
        df3_RCH2_XY.insert(3,'RCH_X',RCH_Xarrlist)
        df3_RCH2_XY.insert(4,'RCH_Y',RCH_Yarrlist)
        
        df3_RCH2_XY = df3_RCH2_XY[df3_RCH2_XY.columns[3:6]]
        
        #删除上次生成的文件,避免出现无法修改数据错误
        RCH_PATH = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH2-Mapdata.txt'
        RCH_PATH1 = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH2-Mapdata.xls'
        if os.path.exists(RCH_PATH):
            os.remove(RCH_PATH)
        elif os.path.exists(RCH_PATH1):
            os.remove(RCH_PATH1)
        else:
            pass
            
        #存储为txt文件
        with open(RCH_PATH,'a+',encoding='utf-8') as f1:
            for line in df3_RCH2_XY.values:
                f1.write((str(line[0])+'\t'+str(line[1])+'\t'+str(line[2])+'\n'))
        workbook1 = xlwt.Workbook()
        worksheet1 = workbook1.add_sheet('RCH2-Map')
        
        #txt文件地址
        filein  = RCH_PATH
        readfile  = open(filein,'r')
        for line in readfile:
            word = line.split()
            if len(word) == 3:
                y=int(word[1])
                x=int(word[0])
                RCH = word[2]
                if 1:
                    if (x == 5):    
                        worksheet1.write(x+7,y,RCH,stylePaleBlueBkg)
                    elif (x == 6):
                        worksheet1.write(x+5,y,RCH,styleLightGreenBkg)
                    elif (x == 7):
                        worksheet1.write(x+3,y,RCH,styleSkyBlueBkg)
                    elif (x == 8):
                        worksheet1.write(x+1,y,RCH,styleYellowBkg)
                    elif (x == 9):
                        worksheet1.write(x-1,y,RCH,styleLightBlueBkg)
                    elif (x == 10):
                        worksheet1.write(x-3,y,RCH,styleIceBlueBkg)
                    elif (x == 11):
                        worksheet1.write(x-5,y,RCH,styleOrangeBkg)
                    elif (x == 12):
                        worksheet1.write(x-7,y,RCH,stylePaleBlueBkg)
                    else:
                        worksheet1.write(x,y,RCH,styleOrangeBkg)
        for ix in range(256):
            worksheet1.col(ix).width =  256 * (5 + 1)
            
        workbook1.save(RCH_PATH1)
        print("晶圆RCH2电阻分布MAP完成!")
        
    def Get_XYRCH3New_Map(self):
        global df3_cnt, df3_RCH3_XY
        
        #颜色定义
        styleLightGreenBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')
        styleYellowBkg = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        styleLightBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;')
        styleIceBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue;')
        styleSkyBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;')
        styleOrangeBkg = xlwt.easyxf('pattern: pattern solid, fore_colour orange;')
        stylePaleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;')
        
        if df3_cnt == 46:
            arr1 = list([4 for x in range(46)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH3_XY['X'].tolist()
            RCH_Y = df3_RCH3_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(46):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(46):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        elif df3_cnt == 100:
            arr1 = list([8 for x in range(100)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH3_XY['X'].tolist()
            RCH_Y = df3_RCH3_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(100):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(100):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        else:
            pass
            
        df3_RCH3_XY.insert(3,'RCH_X',RCH_Xarrlist)
        df3_RCH3_XY.insert(4,'RCH_Y',RCH_Yarrlist)
        
        df3_RCH3_XY = df3_RCH3_XY[df3_RCH3_XY.columns[3:6]]
        
        #删除上次生成的文件,避免出现无法修改数据错误
        RCH_PATH = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH3-Mapdata.txt'
        RCH_PATH1 = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH3-Mapdata.xls'
        if os.path.exists(RCH_PATH):
            os.remove(RCH_PATH)
        elif os.path.exists(RCH_PATH1):
            os.remove(RCH_PATH1)
        else:
            pass
            
        #存储为txt文件
        with open(RCH_PATH,'a+',encoding='utf-8') as f1:
            for line in df3_RCH3_XY.values:
                f1.write((str(line[0])+'\t'+str(line[1])+'\t'+str(line[2])+'\n'))
        workbook1 = xlwt.Workbook()
        worksheet1 = workbook1.add_sheet('RCH3-Map')
        
        #txt文件地址
        filein  = RCH_PATH
        readfile  = open(filein,'r')
        for line in readfile:
            word = line.split()
            if len(word) == 3:
                y=int(word[1])
                x=int(word[0])
                RCH = word[2]
                if 1:
                    if (x == 5):    
                        worksheet1.write(x+7,y,RCH,stylePaleBlueBkg)
                    elif (x == 6):
                        worksheet1.write(x+5,y,RCH,styleLightGreenBkg)
                    elif (x == 7):
                        worksheet1.write(x+3,y,RCH,styleSkyBlueBkg)
                    elif (x == 8):
                        worksheet1.write(x+1,y,RCH,styleYellowBkg)
                    elif (x == 9):
                        worksheet1.write(x-1,y,RCH,styleLightBlueBkg)
                    elif (x == 10):
                        worksheet1.write(x-3,y,RCH,styleIceBlueBkg)
                    elif (x == 11):
                        worksheet1.write(x-5,y,RCH,styleOrangeBkg)
                    elif (x == 12):
                        worksheet1.write(x-7,y,RCH,stylePaleBlueBkg)
                    else:
                        worksheet1.write(x,y,RCH,styleOrangeBkg)
        for ix in range(256):
            worksheet1.col(ix).width =  256 * (5 + 1)
            
        workbook1.save(RCH_PATH1)
        print("晶圆RCH3电阻分布MAP完成!")
        
    def Get_XYRCH4New_Map(self):
        global df3_cnt, df3_RCH4_XY
        
        #颜色定义
        styleLightGreenBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')
        styleYellowBkg = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        styleLightBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;')
        styleIceBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue;')
        styleSkyBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;')
        styleOrangeBkg = xlwt.easyxf('pattern: pattern solid, fore_colour orange;')
        stylePaleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;')
        
        if df3_cnt == 46:
            arr1 = list([4 for x in range(46)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH4_XY['X'].tolist()
            RCH_Y = df3_RCH4_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(46):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(46):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        elif df3_cnt == 100:
            arr1 = list([8 for x in range(100)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH4_XY['X'].tolist()
            RCH_Y = df3_RCH4_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(100):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(100):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        else:
            pass
            
        df3_RCH4_XY.insert(3,'RCH_X',RCH_Xarrlist)
        df3_RCH4_XY.insert(4,'RCH_Y',RCH_Yarrlist)
        
        df3_RCH4_XY = df3_RCH4_XY[df3_RCH4_XY.columns[3:6]]
        
        #删除上次生成的文件,避免出现无法修改数据错误
        RCH_PATH = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH4-Mapdata.txt'
        RCH_PATH1 = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH4-Mapdata.xls'
        if os.path.exists(RCH_PATH):
            os.remove(RCH_PATH)
        elif os.path.exists(RCH_PATH1):
            os.remove(RCH_PATH1)
        else:
            pass
            
        #存储为txt文件
        with open(RCH_PATH,'a+',encoding='utf-8') as f1:
            for line in df3_RCH4_XY.values:
                f1.write((str(line[0])+'\t'+str(line[1])+'\t'+str(line[2])+'\n'))
        workbook1 = xlwt.Workbook()
        worksheet1 = workbook1.add_sheet('RCH4-Map')
        
        #txt文件地址
        filein  = RCH_PATH
        readfile  = open(filein,'r')
        for line in readfile:
            word = line.split()
            if len(word) == 3:
                y=int(word[1])
                x=int(word[0])
                RCH = word[2]
                if 1:
                    if (x == 5):    
                        worksheet1.write(x+7,y,RCH,stylePaleBlueBkg)
                    elif (x == 6):
                        worksheet1.write(x+5,y,RCH,styleLightGreenBkg)
                    elif (x == 7):
                        worksheet1.write(x+3,y,RCH,styleSkyBlueBkg)
                    elif (x == 8):
                        worksheet1.write(x+1,y,RCH,styleYellowBkg)
                    elif (x == 9):
                        worksheet1.write(x-1,y,RCH,styleLightBlueBkg)
                    elif (x == 10):
                        worksheet1.write(x-3,y,RCH,styleIceBlueBkg)
                    elif (x == 11):
                        worksheet1.write(x-5,y,RCH,styleOrangeBkg)
                    elif (x == 12):
                        worksheet1.write(x-7,y,RCH,stylePaleBlueBkg)
                    else:
                        worksheet1.write(x,y,RCH,styleOrangeBkg)
        for ix in range(256):
            worksheet1.col(ix).width =  256 * (5 + 1)
            
        workbook1.save(RCH_PATH1)
        print("晶圆RCH4电阻分布MAP完成!")
        
    def Get_XYRCH5New_Map(self):
        global df3_cnt, df3_RCH5_XY
        
        #颜色定义
        styleLightGreenBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')
        styleYellowBkg = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        styleLightBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;')
        styleIceBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue;')
        styleSkyBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;')
        styleOrangeBkg = xlwt.easyxf('pattern: pattern solid, fore_colour orange;')
        stylePaleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;')
        
        if df3_cnt == 46:
            arr1 = list([4 for x in range(46)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH5_XY['X'].tolist()
            RCH_Y = df3_RCH5_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(46):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(46):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        elif df3_cnt == 100:
            arr1 = list([8 for x in range(100)])  #产生值为4的list列表
            RCH_X = []
            RCH_Y = []
        
            RCH_X = df3_RCH5_XY['X'].tolist()
            RCH_Y = df3_RCH5_XY['Y'].tolist()
            RCH_X = list(map(int,RCH_X))
            RCH_Y = list(map(int,RCH_Y))
            RCH_Xlist = []
            RCH_Ylist = []
            RCH_Xarrlist = []
            RCH_Yarrlist = []
            for i in range(100):
                RCH_Xlist = RCH_X[i] + arr1[i]
                RCH_Xlist = str(RCH_Xlist)  
                RCH_Xarrlist.append(RCH_Xlist)
                
            for j in range(100):
                RCH_Ylist = RCH_Y[j] + arr1[j]
                RCH_Ylist = str(RCH_Ylist)
                RCH_Yarrlist.append(RCH_Ylist)
                
        else:
            pass
            
        df3_RCH5_XY.insert(3,'RCH_X',RCH_Xarrlist)
        df3_RCH5_XY.insert(4,'RCH_Y',RCH_Yarrlist)
        
        df3_RCH5_XY = df3_RCH5_XY[df3_RCH5_XY.columns[3:6]]
        
        #删除上次生成的文件,避免出现无法修改数据错误
        RCH_PATH = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH5-Mapdata.txt'
        RCH_PATH1 = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\RCH5-Mapdata.xls'
        if os.path.exists(RCH_PATH):
            os.remove(RCH_PATH)
        elif os.path.exists(RCH_PATH1):
            os.remove(RCH_PATH1)
        else:
            pass
            
        #存储为txt文件
        with open(RCH_PATH,'a+',encoding='utf-8') as f1:
            for line in df3_RCH5_XY.values:
                f1.write((str(line[0])+'\t'+str(line[1])+'\t'+str(line[2])+'\n'))
        workbook1 = xlwt.Workbook()
        worksheet1 = workbook1.add_sheet('RCH5-Map')
        
        #txt文件地址
        filein  = RCH_PATH
        readfile  = open(filein,'r')
        for line in readfile:
            word = line.split()
            if len(word) == 3:
                y=int(word[1])
                x=int(word[0])
                RCH = word[2]
                if 1:
                    if (x == 5):    
                        worksheet1.write(x+7,y,RCH,stylePaleBlueBkg)
                    elif (x == 6):
                        worksheet1.write(x+5,y,RCH,styleLightGreenBkg)
                    elif (x == 7):
                        worksheet1.write(x+3,y,RCH,styleSkyBlueBkg)
                    elif (x == 8):
                        worksheet1.write(x+1,y,RCH,styleYellowBkg)
                    elif (x == 9):
                        worksheet1.write(x-1,y,RCH,styleLightBlueBkg)
                    elif (x == 10):
                        worksheet1.write(x-3,y,RCH,styleIceBlueBkg)
                    elif (x == 11):
                        worksheet1.write(x-5,y,RCH,styleOrangeBkg)
                    elif (x == 12):
                        worksheet1.write(x-7,y,RCH,stylePaleBlueBkg)
                    else:
                        worksheet1.write(x,y,RCH,styleOrangeBkg)
        for ix in range(256):
            worksheet1.col(ix).width =  256 * (5 + 1)
            
        workbook1.save(RCH_PATH1)
        print("晶圆RCH5电阻分布MAP完成!")
    
    @pyqtSlot()
    def on_btn_SlotRes_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
#        print("点击平板区电阻分布按钮")
        self.Get_csv()
        self.Get_New_csv()
        self.Wafer_Dies()
        self.Get_RCHSlot()
        self.Get_RCHx_Data_csv()
        self.Get_RCH_Mean()
        
        self.Get_XYRCH1_Mean_csv()
        self.Get_XYRCH1New_Map()
        
        self.Get_XYRCH2_Mean_csv()
        self.Get_XYRCH2New_Map()
        
        self.Get_XYRCH3_Mean_csv()
        self.Get_XYRCH3New_Map()
        
        self.Get_XYRCH4_Mean_csv()
        self.Get_XYRCH4New_Map()
        
        self.Get_XYRCH5_Mean_csv()
        self.Get_XYRCH5New_Map()
        print("平板区电阻分布分析完成!")
