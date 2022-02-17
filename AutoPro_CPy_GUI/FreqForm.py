# -*- coding: utf-8 -*-

"""
Module implementing FreqForm.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget

from Ui_FreqForm import Ui_Form
import pandas as pd
import os
import xlwt

class FreqForm(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    df1 = []#全局变量定义
    df3 = []
    df3_new = []
    df3_XY = []
    df3_cnt = []
    df3_freq = []
    
    df3_Cap_mean = []
    
    
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget (defaults to None)
        @type QWidget (optional)
        """
        super(FreqForm, self).__init__(parent)
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
            
    '''获取频率测试原始数据'''
    def Get_Freq_raw_data(self):
        global df3, df3_new, df3_freq
        Cap_cols_num = [i for i,x in enumerate(df3.columns) if x.find('FREQ')!=-1]
        Cap_cols = df3.columns[Cap_cols_num]
        
        for Capcols_num in Cap_cols:
            df3_new[Capcols_num] = df3_new[Capcols_num].astype('float64')#类型转换
            
        df3_freq = df3_new[Cap_cols]
        
    '''获取频率平均值'''
    def Get_Freq_mean(self):
        global df3_freq, df3_Cap_mean
        df3_freq_cols1 = df3_freq.columns
        df3_freq_index1 = 0
        #数据阈值定义
        df3_Freq_Cap = 10.0
#        df3_Cap_mean = [] #@Date:2021/04/27
        
        #单颗芯片52通道频率求和,变量定义
        Cap_sum1 = 0;Cap_sum2 = 0;Cap_sum3 = 0;Cap_sum4 = 0;Cap_sum5 = 0;Cap_sum6 = 0;   Cap_sum7 = 0;    Cap_sum8 = 0;
        Cap_sum9 = 0;Cap_sum10 = 0;Cap_sum11 = 0;Cap_sum12 = 0;Cap_sum13 = 0;Cap_sum14 = 0;Cap_sum15 = 0; Cap_sum16 = 0;
        Cap_sum17 = 0;Cap_sum18 = 0;Cap_sum19 = 0;Cap_sum20 = 0;Cap_sum21 = 0;Cap_sum22 = 0;Cap_sum23 = 0;Cap_sum24 = 0;
        Cap_sum25 = 0;Cap_sum26 = 0;Cap_sum27 = 0;Cap_sum28 = 0;Cap_sum29 = 0;Cap_sum30 = 0;Cap_sum31 = 0;Cap_sum32 = 0;
        Cap_sum33 = 0;Cap_sum34 = 0;Cap_sum35 = 0;Cap_sum36 = 0;Cap_sum37 = 0;Cap_sum38 = 0;Cap_sum39 = 0;Cap_sum40 = 0;
        Cap_sum41 = 0;Cap_sum42 = 0;Cap_sum43 = 0;Cap_sum44 = 0;Cap_sum45 = 0;Cap_sum46 = 0;Cap_sum47 = 0;Cap_sum48 = 0;
        Cap_sum49 = 0;Cap_sum50 = 0;Cap_sum51 = 0;Cap_sum52 = 0;
        Cap_sum53 = 0;Cap_sum54 = 0;Cap_sum55 = 0;Cap_sum56 = 0;Cap_sum57 = 0;Cap_sum58 = 0;Cap_sum59 = 0;Cap_sum60 = 0;
        Cap_sum61 = 0;Cap_sum62 = 0;Cap_sum63 = 0;Cap_sum64 = 0;Cap_sum65 = 0;Cap_sum66 = 0;Cap_sum67 = 0;Cap_sum68 = 0;
        Cap_sum69 = 0;Cap_sum70 = 0;Cap_sum71 = 0;Cap_sum72 = 0;Cap_sum73 = 0;Cap_sum74 = 0;Cap_sum75 = 0;Cap_sum76 = 0;
        Cap_sum77 = 0;Cap_sum78 = 0;Cap_sum79 = 0;Cap_sum80 = 0;Cap_sum81 = 0;Cap_sum82 = 0;Cap_sum83 = 0;Cap_sum84 = 0;
        Cap_sum85 = 0;Cap_sum86 = 0;Cap_sum87 = 0;Cap_sum88 = 0;Cap_sum89 = 0;Cap_sum90 = 0;Cap_sum91 = 0;Cap_sum92 = 0;
        Cap_sum93 = 0;Cap_sum94 = 0;Cap_sum95 = 0;Cap_sum96 = 0;Cap_sum97 = 0;Cap_sum98 = 0;Cap_sum99 = 0;Cap_sum100 = 0;
        
        #有效频率值计数，变量定义
        Cap_cnt1 = 0;Cap_cnt2 = 0;Cap_cnt3 = 0;Cap_cnt4 = 0;Cap_cnt5 = 0;Cap_cnt6 = 0;Cap_cnt7 = 0;Cap_cnt8 = 0;
        Cap_cnt9 = 0;Cap_cnt10 = 0;Cap_cnt11 = 0;Cap_cnt12 = 0;Cap_cnt13 = 0;Cap_cnt14 = 0;Cap_cnt15 = 0;Cap_cnt16 = 0;
        Cap_cnt17 = 0;Cap_cnt18 = 0;Cap_cnt19 = 0;Cap_cnt20 = 0;Cap_cnt21 = 0;Cap_cnt22 = 0;Cap_cnt23 = 0;Cap_cnt24 = 0;
        Cap_cnt25 = 0;Cap_cnt26 = 0;Cap_cnt27 = 0;Cap_cnt28 = 0;Cap_cnt29 = 0;Cap_cnt30 = 0;Cap_cnt31 = 0;Cap_cnt32 = 0;
        Cap_cnt33 = 0;Cap_cnt34 = 0;Cap_cnt35 = 0;Cap_cnt36 = 0;Cap_cnt37 = 0;Cap_cnt38 = 0;Cap_cnt39 = 0;Cap_cnt40 = 0;
        Cap_cnt41 = 0;Cap_cnt42 = 0;Cap_cnt43 = 0;Cap_cnt44 = 0;Cap_cnt45 = 0;Cap_cnt46 = 0;Cap_cnt47 = 0;Cap_cnt48 = 0;
        Cap_cnt49 = 0;Cap_cnt50 = 0;Cap_cnt51 = 0;Cap_cnt52 = 0;
        Cap_cnt53 = 0;Cap_cnt54 = 0;Cap_cnt55 = 0;Cap_cnt56 = 0;Cap_cnt57 = 0;Cap_cnt58 = 0;Cap_cnt59 = 0;Cap_cnt60 = 0;
        Cap_cnt61 = 0;Cap_cnt62 = 0;Cap_cnt63 = 0;Cap_cnt64 = 0;Cap_cnt65 = 0;Cap_cnt66 = 0;Cap_cnt67 = 0;Cap_cnt68 = 0;
        Cap_cnt69 = 0;Cap_cnt70 = 0;Cap_cnt71 = 0;Cap_cnt72 = 0;Cap_cnt73 = 0;Cap_cnt74 = 0;Cap_cnt75 = 0;Cap_cnt76 = 0;
        Cap_cnt77 = 0;Cap_cnt78 = 0;Cap_cnt79 = 0;Cap_cnt80 = 0;Cap_cnt81 = 0;Cap_cnt82 = 0;Cap_cnt83 = 0;Cap_cnt84 = 0;
        Cap_cnt85 = 0;Cap_cnt86 = 0;Cap_cnt87 = 0;Cap_cnt88 = 0;Cap_cnt89 = 0;Cap_cnt90 = 0;Cap_cnt91 = 0;Cap_cnt92 = 0;
        Cap_cnt93 = 0;Cap_cnt94 = 0;Cap_cnt95 = 0;Cap_cnt96 = 0;Cap_cnt97 = 0;Cap_cnt98 = 0;Cap_cnt99 = 0;Cap_cnt100 = 0;
        
        #有效频率值求解平均值,变量定义
        Cap_mean1 = 0;Cap_mean2 = 0;Cap_mean3 = 0;Cap_mean4 = 0;Cap_mean5 = 0;Cap_mean6 = 0;Cap_mean7 = 0;Cap_mean8 = 0;
        Cap_mean9 = 0;Cap_mean10 = 0;Cap_mean11 = 0;Cap_mean12 = 0;Cap_mean13 = 0;Cap_mean14 = 0;Cap_mean15 = 0;Cap_mean16 = 0;
        Cap_mean17 = 0;Cap_mean18 = 0;Cap_mean19 = 0;Cap_mean20 = 0;Cap_mean21 = 0;Cap_mean22 = 0;Cap_mean23 = 0;Cap_mean24 = 0;
        Cap_mean25 = 0;Cap_mean26 = 0;Cap_mean27 = 0;Cap_mean28 = 0;Cap_mean29 = 0;Cap_mean30 = 0;Cap_mean31 = 0;Cap_mean32 = 0;
        Cap_mean33 = 0;Cap_mean34 = 0;Cap_mean35 = 0;Cap_mean36 = 0;Cap_mean37 = 0;Cap_mean38 = 0;Cap_mean39 = 0;Cap_mean40 = 0;
        Cap_mean41 = 0;Cap_mean42 = 0;Cap_mean43 = 0;Cap_mean44 = 0;Cap_mean45 = 0;Cap_mean46 = 0;Cap_mean47 = 0;Cap_mean48 = 0;
        Cap_mean49 = 0;Cap_mean50 = 0;Cap_mean51 = 0;Cap_mean52 = 0;
        Cap_mean53 = 0;Cap_mean54 = 0;Cap_mean55 = 0;Cap_mean56 = 0;Cap_mean57 = 0;Cap_mean58 = 0;Cap_mean59 = 0;Cap_mean60 = 0;
        Cap_mean61 = 0;Cap_mean62 = 0;Cap_mean63 = 0;Cap_mean64 = 0;Cap_mean65 = 0;Cap_mean66 = 0;Cap_mean67 = 0;Cap_mean68 = 0;
        Cap_mean69 = 0;Cap_mean70 = 0;Cap_mean71 = 0;Cap_mean72 = 0;Cap_mean73 = 0;Cap_mean74 = 0;Cap_mean75 = 0;Cap_mean76 = 0;
        Cap_mean77 = 0;Cap_mean78 = 0;Cap_mean79 = 0;Cap_mean80 = 0;Cap_mean81 = 0;Cap_mean82 = 0;Cap_mean83 = 0;Cap_mean84 = 0;
        Cap_mean85 = 0;Cap_mean86 = 0;Cap_mean87 = 0;Cap_mean88 = 0;Cap_mean89 = 0;Cap_mean90 = 0;Cap_mean91 = 0;Cap_mean92 = 0;
        Cap_mean93 = 0;Cap_mean94 = 0;Cap_mean95 = 0;Cap_mean96 = 0;Cap_mean97 = 0;Cap_mean98 = 0;Cap_mean99 = 0;Cap_mean100 = 0;
        
        for df3_freq_num in df3_freq_cols1:
            df3_freq_Cap1 = df3_freq.iloc[df3_freq_index1 + 0][df3_freq_num]
            df3_freq_Cap2 = df3_freq.iloc[df3_freq_index1 + 1][df3_freq_num]
            df3_freq_Cap3 = df3_freq.iloc[df3_freq_index1 + 2][df3_freq_num]
            df3_freq_Cap4 = df3_freq.iloc[df3_freq_index1 + 3][df3_freq_num]
            df3_freq_Cap5 = df3_freq.iloc[df3_freq_index1 + 4][df3_freq_num]
            df3_freq_Cap6 = df3_freq.iloc[df3_freq_index1 + 5][df3_freq_num]
            df3_freq_Cap7 = df3_freq.iloc[df3_freq_index1 + 6][df3_freq_num]
            df3_freq_Cap8 = df3_freq.iloc[df3_freq_index1 + 7][df3_freq_num]
            df3_freq_Cap9 = df3_freq.iloc[df3_freq_index1 + 8][df3_freq_num]
            df3_freq_Cap10 = df3_freq.iloc[df3_freq_index1 + 9][df3_freq_num]
            df3_freq_Cap11 = df3_freq.iloc[df3_freq_index1 + 10][df3_freq_num]
            df3_freq_Cap12 = df3_freq.iloc[df3_freq_index1 + 11][df3_freq_num]
            df3_freq_Cap13 = df3_freq.iloc[df3_freq_index1 + 12][df3_freq_num]
            df3_freq_Cap14 = df3_freq.iloc[df3_freq_index1 + 13][df3_freq_num]
            df3_freq_Cap15 = df3_freq.iloc[df3_freq_index1 + 14][df3_freq_num]
            df3_freq_Cap16 = df3_freq.iloc[df3_freq_index1 + 15][df3_freq_num]
            df3_freq_Cap17 = df3_freq.iloc[df3_freq_index1 + 16][df3_freq_num]
            df3_freq_Cap18 = df3_freq.iloc[df3_freq_index1 + 17][df3_freq_num]
            df3_freq_Cap19 = df3_freq.iloc[df3_freq_index1 + 18][df3_freq_num]
            df3_freq_Cap20 = df3_freq.iloc[df3_freq_index1 + 19][df3_freq_num]
            df3_freq_Cap21 = df3_freq.iloc[df3_freq_index1 + 20][df3_freq_num]
            df3_freq_Cap22 = df3_freq.iloc[df3_freq_index1 + 21][df3_freq_num]
            df3_freq_Cap23 = df3_freq.iloc[df3_freq_index1 + 22][df3_freq_num]
            df3_freq_Cap24 = df3_freq.iloc[df3_freq_index1 + 23][df3_freq_num]
            df3_freq_Cap25 = df3_freq.iloc[df3_freq_index1 + 24][df3_freq_num]
            df3_freq_Cap26 = df3_freq.iloc[df3_freq_index1 + 25][df3_freq_num]
            df3_freq_Cap27 = df3_freq.iloc[df3_freq_index1 + 26][df3_freq_num]
            df3_freq_Cap28 = df3_freq.iloc[df3_freq_index1 + 27][df3_freq_num]
            df3_freq_Cap29 = df3_freq.iloc[df3_freq_index1 + 28][df3_freq_num]
            df3_freq_Cap30 = df3_freq.iloc[df3_freq_index1 + 29][df3_freq_num]
            df3_freq_Cap31 = df3_freq.iloc[df3_freq_index1 + 30][df3_freq_num]
            df3_freq_Cap32 = df3_freq.iloc[df3_freq_index1 + 31][df3_freq_num]
            df3_freq_Cap33 = df3_freq.iloc[df3_freq_index1 + 32][df3_freq_num];
            df3_freq_Cap34 = df3_freq.iloc[df3_freq_index1 + 33][df3_freq_num]
            df3_freq_Cap35 = df3_freq.iloc[df3_freq_index1 + 34][df3_freq_num];
            df3_freq_Cap36 = df3_freq.iloc[df3_freq_index1 + 35][df3_freq_num];
            df3_freq_Cap37 = df3_freq.iloc[df3_freq_index1 + 36][df3_freq_num];
            df3_freq_Cap38 = df3_freq.iloc[df3_freq_index1 + 37][df3_freq_num];
            df3_freq_Cap39 = df3_freq.iloc[df3_freq_index1 + 38][df3_freq_num];
            df3_freq_Cap40 = df3_freq.iloc[df3_freq_index1 + 39][df3_freq_num];
            df3_freq_Cap41 = df3_freq.iloc[df3_freq_index1 + 40][df3_freq_num];
            df3_freq_Cap42 = df3_freq.iloc[df3_freq_index1 + 41][df3_freq_num];
            df3_freq_Cap43 = df3_freq.iloc[df3_freq_index1 + 42][df3_freq_num];
            df3_freq_Cap44 = df3_freq.iloc[df3_freq_index1 + 43][df3_freq_num];
            df3_freq_Cap45 = df3_freq.iloc[df3_freq_index1 + 44][df3_freq_num];
            df3_freq_Cap46 = df3_freq.iloc[df3_freq_index1 + 45][df3_freq_num];
            df3_freq_Cap47 = df3_freq.iloc[df3_freq_index1 + 46][df3_freq_num];
            df3_freq_Cap48 = df3_freq.iloc[df3_freq_index1 + 47][df3_freq_num];
            df3_freq_Cap49 = df3_freq.iloc[df3_freq_index1 + 48][df3_freq_num];
            df3_freq_Cap50 = df3_freq.iloc[df3_freq_index1 + 49][df3_freq_num];
            df3_freq_Cap51 = df3_freq.iloc[df3_freq_index1 + 50][df3_freq_num];
            df3_freq_Cap52 = df3_freq.iloc[df3_freq_index1 + 51][df3_freq_num];
            df3_freq_Cap53 = df3_freq.iloc[df3_freq_index1 + 52][df3_freq_num];
            df3_freq_Cap54 = df3_freq.iloc[df3_freq_index1 + 53][df3_freq_num];
            df3_freq_Cap55 = df3_freq.iloc[df3_freq_index1 + 54][df3_freq_num];
            df3_freq_Cap56 = df3_freq.iloc[df3_freq_index1 + 55][df3_freq_num];
            df3_freq_Cap57 = df3_freq.iloc[df3_freq_index1 + 56][df3_freq_num];
            df3_freq_Cap58 = df3_freq.iloc[df3_freq_index1 + 57][df3_freq_num];
            df3_freq_Cap59 = df3_freq.iloc[df3_freq_index1 + 58][df3_freq_num];
            df3_freq_Cap60 = df3_freq.iloc[df3_freq_index1 + 59][df3_freq_num];
            df3_freq_Cap61 = df3_freq.iloc[df3_freq_index1 + 60][df3_freq_num];
            df3_freq_Cap62 = df3_freq.iloc[df3_freq_index1 + 61][df3_freq_num];
            df3_freq_Cap63 = df3_freq.iloc[df3_freq_index1 + 62][df3_freq_num];
            df3_freq_Cap64 = df3_freq.iloc[df3_freq_index1 + 63][df3_freq_num];
            df3_freq_Cap65 = df3_freq.iloc[df3_freq_index1 + 64][df3_freq_num]
            df3_freq_Cap66 = df3_freq.iloc[df3_freq_index1 + 65][df3_freq_num]
            df3_freq_Cap67 = df3_freq.iloc[df3_freq_index1 + 66][df3_freq_num]
            df3_freq_Cap68 = df3_freq.iloc[df3_freq_index1 + 67][df3_freq_num]
            df3_freq_Cap69 = df3_freq.iloc[df3_freq_index1 + 68][df3_freq_num]
            df3_freq_Cap70 = df3_freq.iloc[df3_freq_index1 + 69][df3_freq_num]
            df3_freq_Cap71 = df3_freq.iloc[df3_freq_index1 + 70][df3_freq_num]
            df3_freq_Cap72 = df3_freq.iloc[df3_freq_index1 + 71][df3_freq_num]
            df3_freq_Cap73 = df3_freq.iloc[df3_freq_index1 + 72][df3_freq_num]
            df3_freq_Cap74 = df3_freq.iloc[df3_freq_index1 + 73][df3_freq_num]
            df3_freq_Cap75 = df3_freq.iloc[df3_freq_index1 + 74][df3_freq_num]
            df3_freq_Cap76 = df3_freq.iloc[df3_freq_index1 + 75][df3_freq_num]
            df3_freq_Cap77 = df3_freq.iloc[df3_freq_index1 + 76][df3_freq_num]
            df3_freq_Cap78 = df3_freq.iloc[df3_freq_index1 + 77][df3_freq_num]
            df3_freq_Cap79 = df3_freq.iloc[df3_freq_index1 + 78][df3_freq_num]
            df3_freq_Cap80 = df3_freq.iloc[df3_freq_index1 + 79][df3_freq_num]
            df3_freq_Cap81 = df3_freq.iloc[df3_freq_index1 + 80][df3_freq_num]
            df3_freq_Cap82 = df3_freq.iloc[df3_freq_index1 + 81][df3_freq_num]
            df3_freq_Cap83 = df3_freq.iloc[df3_freq_index1 + 82][df3_freq_num]
            df3_freq_Cap84 = df3_freq.iloc[df3_freq_index1 + 83][df3_freq_num]
            df3_freq_Cap85 = df3_freq.iloc[df3_freq_index1 + 84][df3_freq_num]
            df3_freq_Cap86 = df3_freq.iloc[df3_freq_index1 + 85][df3_freq_num]
            df3_freq_Cap87 = df3_freq.iloc[df3_freq_index1 + 86][df3_freq_num]
            df3_freq_Cap88 = df3_freq.iloc[df3_freq_index1 + 87][df3_freq_num]
            df3_freq_Cap89 = df3_freq.iloc[df3_freq_index1 + 88][df3_freq_num]
            df3_freq_Cap90 = df3_freq.iloc[df3_freq_index1 + 89][df3_freq_num]
            df3_freq_Cap91 = df3_freq.iloc[df3_freq_index1 + 90][df3_freq_num]
            df3_freq_Cap92 = df3_freq.iloc[df3_freq_index1 + 91][df3_freq_num]
            df3_freq_Cap93 = df3_freq.iloc[df3_freq_index1 + 92][df3_freq_num]
            df3_freq_Cap94 = df3_freq.iloc[df3_freq_index1 + 93][df3_freq_num]
            df3_freq_Cap95 = df3_freq.iloc[df3_freq_index1 + 94][df3_freq_num]
            df3_freq_Cap96 = df3_freq.iloc[df3_freq_index1 + 95][df3_freq_num]
            df3_freq_Cap97 = df3_freq.iloc[df3_freq_index1 + 96][df3_freq_num]
            df3_freq_Cap98 = df3_freq.iloc[df3_freq_index1 + 97][df3_freq_num]
            df3_freq_Cap99 = df3_freq.iloc[df3_freq_index1 + 98][df3_freq_num]
            df3_freq_Cap100 = df3_freq.iloc[df3_freq_index1 + 99][df3_freq_num]
            if df3_cnt == 46:
                if df3_freq_Cap1 > df3_Freq_Cap + 0:
                    Cap_sum1 += df3_freq_Cap1
                    Cap_cnt1 += 1
                    Cap_mean1 = Cap_sum1 / Cap_cnt1
                if df3_freq_Cap2 > df3_Freq_Cap + 0:
                    Cap_sum2 += df3_freq_Cap2
                    Cap_cnt2 += 1
                    Cap_mean2 = Cap_sum2 / Cap_cnt2
                if df3_freq_Cap3 > df3_Freq_Cap:
                    Cap_sum3 += df3_freq_Cap3
                    Cap_cnt3 += 1
                    Cap_mean3 = Cap_sum3 / Cap_cnt3
                if df3_freq_Cap4 > df3_Freq_Cap:
                    Cap_sum4 += df3_freq_Cap4
                    Cap_cnt4 += 1
                    Cap_mean4 = Cap_sum4 / Cap_cnt4
                if df3_freq_Cap5 > df3_Freq_Cap:
                    Cap_sum5 += df3_freq_Cap5
                    Cap_cnt5 += 1
                    Cap_mean5 = Cap_sum5 / Cap_cnt5
                if df3_freq_Cap6 > df3_Freq_Cap:
                    Cap_sum6 += df3_freq_Cap6
                    Cap_cnt6 += 1
                    Cap_mean6 = Cap_sum6 / Cap_cnt6
                if df3_freq_Cap7 > df3_Freq_Cap:
                    Cap_sum7 += df3_freq_Cap7
                    Cap_cnt7 += 1
                    Cap_mean7 = Cap_sum7 / Cap_cnt7
                if df3_freq_Cap8 > df3_Freq_Cap:
                    Cap_sum8 += df3_freq_Cap8
                    Cap_cnt8 += 1
                    Cap_mean8 = Cap_sum8 / Cap_cnt8
                if df3_freq_Cap9 > df3_Freq_Cap:
                    Cap_sum9 += df3_freq_Cap9
                    Cap_cnt9 += 1
                    Cap_mean9 = Cap_sum9 / Cap_cnt9
                if df3_freq_Cap10 > df3_Freq_Cap:
                    Cap_sum10 += df3_freq_Cap10
                    Cap_cnt10 += 1
                    Cap_mean10 = Cap_sum10 / Cap_cnt10
                if df3_freq_Cap11 > df3_Freq_Cap:
                    Cap_sum11 += df3_freq_Cap11
                    Cap_cnt11 += 1
                    Cap_mean11 = Cap_sum11 / Cap_cnt11
                if df3_freq_Cap12 > df3_Freq_Cap:
                    Cap_sum12 += df3_freq_Cap12
                    Cap_cnt12 += 1
                    Cap_mean12 = Cap_sum12 / Cap_cnt12
                if df3_freq_Cap13 > df3_Freq_Cap:
                    Cap_sum13 += df3_freq_Cap13
                    Cap_cnt13 += 1
                    Cap_mean13 = Cap_sum13 / Cap_cnt13
                if df3_freq_Cap14 > df3_Freq_Cap:
                    Cap_sum14 += df3_freq_Cap14
                    Cap_cnt14 += 1
                    Cap_mean14 = Cap_sum14 / Cap_cnt14
                if df3_freq_Cap15 > df3_Freq_Cap:
                    Cap_sum15 += df3_freq_Cap15
                    Cap_cnt15 += 1
                    Cap_mean15 = Cap_sum15 / Cap_cnt15
                if df3_freq_Cap16 > df3_Freq_Cap:
                    Cap_sum16 += df3_freq_Cap16
                    Cap_cnt16 += 1
                    Cap_mean16 = Cap_sum16 / Cap_cnt16
                if df3_freq_Cap17 > df3_Freq_Cap:
                    Cap_sum17 += df3_freq_Cap17
                    Cap_cnt17 += 1
                    Cap_mean17 = Cap_sum17 / Cap_cnt17
                if df3_freq_Cap18 > df3_Freq_Cap:
                    Cap_sum18 += df3_freq_Cap18
                    Cap_cnt18 += 1
                    Cap_mean18 = Cap_sum18 / Cap_cnt18
                if df3_freq_Cap19 > df3_Freq_Cap:
                    Cap_sum19 += df3_freq_Cap19
                    Cap_cnt19 += 1
                    Cap_mean19 = Cap_sum19 / Cap_cnt19
                if df3_freq_Cap20 > df3_Freq_Cap:
                    Cap_sum20 += df3_freq_Cap20
                    Cap_cnt20 += 1
                    Cap_mean20 = Cap_sum20 / Cap_cnt20
                if df3_freq_Cap21 > df3_Freq_Cap:
                    Cap_sum21 += df3_freq_Cap21
                    Cap_cnt21 += 1
                    Cap_mean21 = Cap_sum21 / Cap_cnt21
                if df3_freq_Cap22 > df3_Freq_Cap:
                    Cap_sum22 += df3_freq_Cap22
                    Cap_cnt22 += 1
                    Cap_mean22 = Cap_sum22 / Cap_cnt22
                if df3_freq_Cap23 > df3_Freq_Cap:
                    Cap_sum23 += df3_freq_Cap23
                    Cap_cnt23 += 1
                    Cap_mean23 = Cap_sum23 / Cap_cnt23
                if df3_freq_Cap24 > df3_Freq_Cap:
                    Cap_sum24 += df3_freq_Cap24
                    Cap_cnt24 += 1
                    Cap_mean24 = Cap_sum24 / Cap_cnt24
                if df3_freq_Cap25 > df3_Freq_Cap:
                    Cap_sum25 += df3_freq_Cap25
                    Cap_cnt25 += 1
                    Cap_mean25 = Cap_sum25 / Cap_cnt25
                if df3_freq_Cap26 > df3_Freq_Cap:
                    Cap_sum26 += df3_freq_Cap26
                    Cap_cnt26 += 1
                    Cap_mean26 = Cap_sum26 / Cap_cnt26
                if df3_freq_Cap27 > df3_Freq_Cap:
                    Cap_sum27 += df3_freq_Cap27
                    Cap_cnt27 += 1
                    Cap_mean27 = Cap_sum27 / Cap_cnt27
                if df3_freq_Cap28 > df3_Freq_Cap:
                    Cap_sum28 += df3_freq_Cap28
                    Cap_cnt28 += 1
                    Cap_mean28 = Cap_sum28 / Cap_cnt28
                if df3_freq_Cap29 > df3_Freq_Cap:
                    Cap_sum29 += df3_freq_Cap29
                    Cap_cnt29 += 1
                    Cap_mean29 = Cap_sum29 / Cap_cnt29
                if df3_freq_Cap30 > df3_Freq_Cap:
                    Cap_sum30 += df3_freq_Cap30
                    Cap_cnt30 += 1
                    Cap_mean30 = Cap_sum30 / Cap_cnt30
                if df3_freq_Cap31 > df3_Freq_Cap:
                    Cap_sum31 += df3_freq_Cap31
                    Cap_cnt31 += 1
                    Cap_mean31 = Cap_sum31 / Cap_cnt31
                if df3_freq_Cap32 > df3_Freq_Cap:
                    Cap_sum32 += df3_freq_Cap32
                    Cap_cnt32 += 1
                    Cap_mean32 = Cap_sum32 / Cap_cnt32
                if df3_freq_Cap33 > df3_Freq_Cap:
                    Cap_sum33 += df3_freq_Cap33
                    Cap_cnt33 += 1
                    Cap_mean33 = Cap_sum33 / Cap_cnt33
                if df3_freq_Cap34 > df3_Freq_Cap:
                    Cap_sum34 += df3_freq_Cap34
                    Cap_cnt34 += 1
                    Cap_mean34 = Cap_sum34 / Cap_cnt34
                if df3_freq_Cap35 > df3_Freq_Cap:
                    Cap_sum35 += df3_freq_Cap35
                    Cap_cnt35 += 1
                    Cap_mean35 = Cap_sum35 / Cap_cnt35
                if df3_freq_Cap36 > df3_Freq_Cap:
                    Cap_sum36 += df3_freq_Cap36
                    Cap_cnt36 += 1
                    Cap_mean36 = Cap_sum36 / Cap_cnt36
                if df3_freq_Cap37 > df3_Freq_Cap:
                    Cap_sum37 += df3_freq_Cap37
                    Cap_cnt37 += 1
                    Cap_mean37 = Cap_sum37 / Cap_cnt37
                if df3_freq_Cap38 > df3_Freq_Cap:
                    Cap_sum38 += df3_freq_Cap38
                    Cap_cnt38 += 1
                    Cap_mean38 = Cap_sum38 / Cap_cnt38
                if df3_freq_Cap39 > df3_Freq_Cap:
                    Cap_sum39 += df3_freq_Cap39
                    Cap_cnt39 += 1
                    Cap_mean39 = Cap_sum39 / Cap_cnt39
                if df3_freq_Cap40 > df3_Freq_Cap:
                    Cap_sum40 += df3_freq_Cap40
                    Cap_cnt40 += 1
                    Cap_mean40 = Cap_sum40 / Cap_cnt40
                if df3_freq_Cap41 > df3_Freq_Cap:
                    Cap_sum41 += df3_freq_Cap41
                    Cap_cnt41 += 1
                    Cap_mean41 = Cap_sum41 / Cap_cnt41
                if df3_freq_Cap42 > df3_Freq_Cap:
                    Cap_sum42 += df3_freq_Cap42
                    Cap_cnt42 += 1
                    Cap_mean42 = Cap_sum42 / Cap_cnt42
                if df3_freq_Cap43 > df3_Freq_Cap:
                    Cap_sum43 += df3_freq_Cap43
                    Cap_cnt43 += 1
                    Cap_mean43 = Cap_sum43 / Cap_cnt43
                if df3_freq_Cap44 > df3_Freq_Cap:
                    Cap_sum44 += df3_freq_Cap44
                    Cap_cnt44 += 1
                    Cap_mean44 = Cap_sum44 / Cap_cnt44
                if df3_freq_Cap45 > df3_Freq_Cap:
                    Cap_sum45 += df3_freq_Cap45
                    Cap_cnt45 += 1
                    Cap_mean45 = Cap_sum45 / Cap_cnt45
                if df3_freq_Cap46 > df3_Freq_Cap:
                    Cap_sum46 += df3_freq_Cap46
                    Cap_cnt46 += 1
                    Cap_mean46 = Cap_sum46 / Cap_cnt46
                else:
                    pass
                    
                df3_Cap_mean = [Cap_mean1,Cap_mean2,Cap_mean3,Cap_mean4,Cap_mean5,Cap_mean6,Cap_mean7,Cap_mean8,Cap_mean9,Cap_mean10,Cap_mean11,\
                     Cap_mean12,Cap_mean13,Cap_mean14,Cap_mean15,Cap_mean16,Cap_mean17,Cap_mean18,Cap_mean19,Cap_mean20,Cap_mean21,Cap_mean22,\
                     Cap_mean23,Cap_mean24,Cap_mean25,Cap_mean26,Cap_mean27,Cap_mean28,Cap_mean29,Cap_mean30,Cap_mean31,Cap_mean32,Cap_mean33,\
                     Cap_mean34,Cap_mean35,Cap_mean36,Cap_mean37,Cap_mean38,Cap_mean39,Cap_mean40,Cap_mean41,Cap_mean42,Cap_mean43,Cap_mean44,\
                     Cap_mean45,Cap_mean46]
                     
            elif df3_cnt == 100:
                if df3_freq_Cap1 > df3_Freq_Cap:
                    Cap_sum1 += df3_freq_Cap1
                    Cap_cnt1 += 1
                    Cap_mean1 = Cap_sum1 / Cap_cnt1
                if df3_freq_Cap2 > df3_Freq_Cap:
                    Cap_sum2 += df3_freq_Cap2
                    Cap_cnt2 += 1
                    Cap_mean2 = Cap_sum2 / Cap_cnt2
                if df3_freq_Cap3 > df3_Freq_Cap:
                    Cap_sum3 += df3_freq_Cap3
                    Cap_cnt3 += 1
                    Cap_mean3 = Cap_sum3 / Cap_cnt3
                if df3_freq_Cap4 > df3_Freq_Cap:
                    Cap_sum4 += df3_freq_Cap4
                    Cap_cnt4 += 1
                    Cap_mean4 = Cap_sum4 / Cap_cnt4
                if df3_freq_Cap5 > df3_Freq_Cap:
                    Cap_sum5 += df3_freq_Cap5
                    Cap_cnt5 += 1
                    Cap_mean5 = Cap_sum5 / Cap_cnt5
                if df3_freq_Cap6 > df3_Freq_Cap:
                    Cap_sum6 += df3_freq_Cap6
                    Cap_cnt6 += 1
                    Cap_mean6 = Cap_sum6 / Cap_cnt6
                if df3_freq_Cap7 > df3_Freq_Cap:
                    Cap_sum7 += df3_freq_Cap7
                    Cap_cnt7 += 1
                    Cap_mean7 = Cap_sum7 / Cap_cnt7
                if df3_freq_Cap8 > df3_Freq_Cap:
                    Cap_sum8 += df3_freq_Cap8
                    Cap_cnt8 += 1
                    Cap_mean8 = Cap_sum8 / Cap_cnt8
                if df3_freq_Cap9 > df3_Freq_Cap:
                    Cap_sum9 += df3_freq_Cap9
                    Cap_cnt9 += 1
                    Cap_mean9 = Cap_sum9 / Cap_cnt9
                if df3_freq_Cap10 > df3_Freq_Cap:
                    Cap_sum10 += df3_freq_Cap10
                    Cap_cnt10 += 1
                    Cap_mean10 = Cap_sum10 / Cap_cnt10
                if df3_freq_Cap11 > df3_Freq_Cap:
                    Cap_sum11 += df3_freq_Cap11
                    Cap_cnt11 += 1
                    Cap_mean11 = Cap_sum11 / Cap_cnt11
                if df3_freq_Cap12 > df3_Freq_Cap:
                    Cap_sum12 += df3_freq_Cap12
                    Cap_cnt12 += 1
                    Cap_mean12 = Cap_sum12 / Cap_cnt12
                if df3_freq_Cap13 > df3_Freq_Cap:
                    Cap_sum13 += df3_freq_Cap13
                    Cap_cnt13 += 1
                    Cap_mean13 = Cap_sum13 / Cap_cnt13
                if df3_freq_Cap14 > df3_Freq_Cap:
                    Cap_sum14 += df3_freq_Cap14
                    Cap_cnt14 += 1
                    Cap_mean14 = Cap_sum14 / Cap_cnt14
                if df3_freq_Cap15 > df3_Freq_Cap:
                    Cap_sum15 += df3_freq_Cap15
                    Cap_cnt15 += 1
                    Cap_mean15 = Cap_sum15 / Cap_cnt15
                if df3_freq_Cap16 > df3_Freq_Cap:
                    Cap_sum16 += df3_freq_Cap16
                    Cap_cnt16 += 1
                    Cap_mean16 = Cap_sum16 / Cap_cnt16
                if df3_freq_Cap17 > df3_Freq_Cap:
                    Cap_sum17 += df3_freq_Cap17
                    Cap_cnt17 += 1
                    Cap_mean17 = Cap_sum17 / Cap_cnt17
                if df3_freq_Cap18 > df3_Freq_Cap:
                    Cap_sum18 += df3_freq_Cap18
                    Cap_cnt18 += 1
                    Cap_mean18 = Cap_sum18 / Cap_cnt18
                if df3_freq_Cap19 > df3_Freq_Cap:
                    Cap_sum19 += df3_freq_Cap19
                    Cap_cnt19 += 1
                    Cap_mean19 = Cap_sum19 / Cap_cnt19
                if df3_freq_Cap20 > df3_Freq_Cap:
                    Cap_sum20 += df3_freq_Cap20
                    Cap_cnt20 += 1
                    Cap_mean20 = Cap_sum20 / Cap_cnt20
                if df3_freq_Cap21 > df3_Freq_Cap:
                    Cap_sum21 += df3_freq_Cap21
                    Cap_cnt21 += 1
                    Cap_mean21 = Cap_sum21 / Cap_cnt21
                if df3_freq_Cap22 > df3_Freq_Cap:
                    Cap_sum22 += df3_freq_Cap22
                    Cap_cnt22 += 1
                    Cap_mean22 = Cap_sum22 / Cap_cnt22
                if df3_freq_Cap23 > df3_Freq_Cap:
                    Cap_sum23 += df3_freq_Cap23
                    Cap_cnt23 += 1
                    Cap_mean23 = Cap_sum23 / Cap_cnt23
                if df3_freq_Cap24 > df3_Freq_Cap:
                    Cap_sum24 += df3_freq_Cap24
                    Cap_cnt24 += 1
                    Cap_mean24 = Cap_sum24 / Cap_cnt24
                if df3_freq_Cap25 > df3_Freq_Cap:
                    Cap_sum25 += df3_freq_Cap25
                    Cap_cnt25 += 1
                    Cap_mean25 = Cap_sum25 / Cap_cnt25
                if df3_freq_Cap26 > df3_Freq_Cap:
                    Cap_sum26 += df3_freq_Cap26
                    Cap_cnt26 += 1
                    Cap_mean26 = Cap_sum26 / Cap_cnt26
                if df3_freq_Cap27 > df3_Freq_Cap:
                    Cap_sum27 += df3_freq_Cap27
                    Cap_cnt27 += 1
                    Cap_mean27 = Cap_sum27 / Cap_cnt27
                if df3_freq_Cap28 > df3_Freq_Cap:
                    Cap_sum28 += df3_freq_Cap28
                    Cap_cnt28 += 1
                    Cap_mean28 = Cap_sum28 / Cap_cnt28
                if df3_freq_Cap29 > df3_Freq_Cap:
                    Cap_sum29 += df3_freq_Cap29
                    Cap_cnt29 += 1
                    Cap_mean29 = Cap_sum29 / Cap_cnt29
                if df3_freq_Cap30 > df3_Freq_Cap:
                    Cap_sum30 += df3_freq_Cap30
                    Cap_cnt30 += 1
                    Cap_mean30 = Cap_sum30 / Cap_cnt30
                if df3_freq_Cap31 > df3_Freq_Cap:
                    Cap_sum31 += df3_freq_Cap31
                    Cap_cnt31 += 1
                    Cap_mean31 = Cap_sum31 / Cap_cnt31
                if df3_freq_Cap32 > df3_Freq_Cap:
                    Cap_sum32 += df3_freq_Cap32
                    Cap_cnt32 += 1
                    Cap_mean32 = Cap_sum32 / Cap_cnt32
                if df3_freq_Cap33 > df3_Freq_Cap:
                    Cap_sum33 += df3_freq_Cap33
                    Cap_cnt33 += 1
                    Cap_mean33 = Cap_sum33 / Cap_cnt33
                if df3_freq_Cap34 > df3_Freq_Cap:
                    Cap_sum34 += df3_freq_Cap34
                    Cap_cnt34 += 1
                    Cap_mean34 = Cap_sum34 / Cap_cnt34
                if df3_freq_Cap35 > df3_Freq_Cap:
                    Cap_sum35 += df3_freq_Cap35
                    Cap_cnt35 += 1
                    Cap_mean35 = Cap_sum35 / Cap_cnt35
                if df3_freq_Cap36 > df3_Freq_Cap:
                    Cap_sum36 += df3_freq_Cap36
                    Cap_cnt36 += 1
                    Cap_mean36 = Cap_sum36 / Cap_cnt36
                if df3_freq_Cap37 > df3_Freq_Cap:
                    Cap_sum37 += df3_freq_Cap37
                    Cap_cnt37 += 1
                    Cap_mean37 = Cap_sum37 / Cap_cnt37
                if df3_freq_Cap38 > df3_Freq_Cap:
                    Cap_sum38 += df3_freq_Cap38
                    Cap_cnt38 += 1
                    Cap_mean38 = Cap_sum38 / Cap_cnt38
                if df3_freq_Cap39 > df3_Freq_Cap:
                    Cap_sum39 += df3_freq_Cap39
                    Cap_cnt39 += 1
                    Cap_mean39 = Cap_sum39 / Cap_cnt39
                if df3_freq_Cap40 > df3_Freq_Cap:
                    Cap_sum40 += df3_freq_Cap40
                    Cap_cnt40 += 1
                    Cap_mean40 = Cap_sum40 / Cap_cnt40
                if df3_freq_Cap41 > df3_Freq_Cap:
                    Cap_sum41 += df3_freq_Cap41
                    Cap_cnt41 += 1
                    Cap_mean41 = Cap_sum41 / Cap_cnt41
                if df3_freq_Cap42 > df3_Freq_Cap:
                    Cap_sum42 += df3_freq_Cap42
                    Cap_cnt42 += 1
                    Cap_mean42 = Cap_sum42 / Cap_cnt42
                if df3_freq_Cap43 > df3_Freq_Cap:
                    Cap_sum43 += df3_freq_Cap43
                    Cap_cnt43 += 1
                    Cap_mean43 = Cap_sum43 / Cap_cnt43
                if df3_freq_Cap44 > df3_Freq_Cap:
                    Cap_sum44 += df3_freq_Cap44
                    Cap_cnt44 += 1
                    Cap_mean44 = Cap_sum44 / Cap_cnt44
                if df3_freq_Cap45 > df3_Freq_Cap:
                    Cap_sum45 += df3_freq_Cap45
                    Cap_cnt45 += 1
                    Cap_mean45 = Cap_sum45 / Cap_cnt45
                if df3_freq_Cap46 > df3_Freq_Cap:
                    Cap_sum46 += df3_freq_Cap46
                    Cap_cnt46 += 1
                    Cap_mean46 = Cap_sum46 / Cap_cnt46
                if df3_freq_Cap47 > df3_Freq_Cap:
                    Cap_sum47 += df3_freq_Cap47
                    Cap_cnt47 += 1
                    Cap_mean47 = Cap_sum47 / Cap_cnt47
                if df3_freq_Cap48 > df3_Freq_Cap:
                    Cap_sum48 += df3_freq_Cap48
                    Cap_cnt48 += 1
                    Cap_mean48 = Cap_sum48 / Cap_cnt48
                if df3_freq_Cap49 > df3_Freq_Cap:
                    Cap_sum49 += df3_freq_Cap49
                    Cap_cnt49 += 1
                    Cap_mean49 = Cap_sum49 / Cap_cnt49
                if df3_freq_Cap50 > df3_Freq_Cap:
                    Cap_sum50 += df3_freq_Cap50
                    Cap_cnt50 += 1
                    Cap_mean50 = Cap_sum50 / Cap_cnt50
                if df3_freq_Cap51 > df3_Freq_Cap:
                    Cap_sum51 += df3_freq_Cap51
                    Cap_cnt51 += 1
                    Cap_mean51 = Cap_sum51 / Cap_cnt51
                if df3_freq_Cap52 > df3_Freq_Cap:
                    Cap_sum52 += df3_freq_Cap52
                    Cap_cnt52 += 1
                    Cap_mean52 = Cap_sum52 / Cap_cnt52
                if df3_freq_Cap53 > df3_Freq_Cap + 5:
                    Cap_sum53 += df3_freq_Cap53
                    Cap_cnt53 += 1
                    Cap_mean53 = Cap_sum53 / Cap_cnt53
                if df3_freq_Cap54 > df3_Freq_Cap + 2:
                    Cap_sum54 += df3_freq_Cap54
                    Cap_cnt54 += 1
                    Cap_mean54 = Cap_sum54 / Cap_cnt54
                if df3_freq_Cap55 > df3_Freq_Cap:
                    Cap_sum55 += df3_freq_Cap55
                    Cap_cnt55 += 1
                    Cap_mean55 = Cap_sum55 / Cap_cnt55
                if df3_freq_Cap56 > df3_Freq_Cap:
                    Cap_sum56 += df3_freq_Cap56
                    Cap_cnt56 += 1
                    Cap_mean56 = Cap_sum56 / Cap_cnt56
                if df3_freq_Cap57 > df3_Freq_Cap:
                    Cap_sum57 += df3_freq_Cap57
                    Cap_cnt57 += 1
                    Cap_mean57 = Cap_sum57 / Cap_cnt57
                if df3_freq_Cap58 > df3_Freq_Cap:
                    Cap_sum58 += df3_freq_Cap58
                    Cap_cnt58 += 1
                    Cap_mean58 = Cap_sum58 / Cap_cnt58
                if df3_freq_Cap59 > df3_Freq_Cap:
                    Cap_sum59 += df3_freq_Cap59
                    Cap_cnt59 += 1
                    Cap_mean59 = Cap_sum59 / Cap_cnt59
                if df3_freq_Cap60 > df3_Freq_Cap:
                    Cap_sum60 += df3_freq_Cap60
                    Cap_cnt60 += 1
                    Cap_mean60 = Cap_sum60 / Cap_cnt60
                if df3_freq_Cap61 > df3_Freq_Cap:
                    Cap_sum61 += df3_freq_Cap61
                    Cap_cnt61 += 1
                    Cap_mean61 = Cap_sum61 / Cap_cnt61
                if df3_freq_Cap62 > df3_Freq_Cap:
                    Cap_sum62 += df3_freq_Cap62
                    Cap_cnt62 += 1
                    Cap_mean62 = Cap_sum62 / Cap_cnt62
                if df3_freq_Cap63 > df3_Freq_Cap:
                    Cap_sum63 += df3_freq_Cap63
                    Cap_cnt63 += 1
                    Cap_mean63 = Cap_sum63 / Cap_cnt63
                if df3_freq_Cap64 > df3_Freq_Cap:
                    Cap_sum64 += df3_freq_Cap64
                    Cap_cnt64 += 1
                    Cap_mean64 = Cap_sum64 / Cap_cnt64
                if df3_freq_Cap65 > df3_Freq_Cap:
                    Cap_sum65 += df3_freq_Cap65
                    Cap_cnt65 += 1
                    Cap_mean65 = Cap_sum65 / Cap_cnt65
                if df3_freq_Cap66 > df3_Freq_Cap:
                    Cap_sum66 += df3_freq_Cap66
                    Cap_cnt66 += 1
                    Cap_mean66 = Cap_sum66 / Cap_cnt66
                if df3_freq_Cap67 > df3_Freq_Cap:
                    Cap_sum67 += df3_freq_Cap67
                    Cap_cnt67 += 1
                    Cap_mean67 = Cap_sum67 / Cap_cnt67
                if df3_freq_Cap68 > df3_Freq_Cap:
                    Cap_sum68 += df3_freq_Cap68
                    Cap_cnt68 += 1
                    Cap_mean68 = Cap_sum68 / Cap_cnt68
                if df3_freq_Cap69 > df3_Freq_Cap:
                    Cap_sum69 += df3_freq_Cap69
                    Cap_cnt69 += 1
                    Cap_mean69 = Cap_sum69 / Cap_cnt69
                if df3_freq_Cap70 > df3_Freq_Cap:
                    Cap_sum70 += df3_freq_Cap70
                    Cap_cnt70 += 1
                    Cap_mean70 = Cap_sum70 / Cap_cnt70
                if df3_freq_Cap71 > df3_Freq_Cap:
                    Cap_sum71 += df3_freq_Cap71
                    Cap_cnt71 += 1
                    Cap_mean71 = Cap_sum71 / Cap_cnt71
                if df3_freq_Cap72 > df3_Freq_Cap:
                    Cap_sum72 += df3_freq_Cap72
                    Cap_cnt72 += 1
                    Cap_mean72 = Cap_sum72 / Cap_cnt72
                if df3_freq_Cap73 > df3_Freq_Cap:
                    Cap_sum73 += df3_freq_Cap73
                    Cap_cnt73 += 1
                    Cap_mean73 = Cap_sum73 / Cap_cnt73
                if df3_freq_Cap74 > df3_Freq_Cap:
                    Cap_sum74 += df3_freq_Cap74
                    Cap_cnt74 += 1
                    Cap_mean74 = Cap_sum74 / Cap_cnt74
                if df3_freq_Cap75 > df3_Freq_Cap:
                    Cap_sum75 += df3_freq_Cap75
                    Cap_cnt75 += 1
                    Cap_mean75 = Cap_sum75 / Cap_cnt75
                if df3_freq_Cap76 > df3_Freq_Cap:
                    Cap_sum76 += df3_freq_Cap76
                    Cap_cnt76 += 1
                    Cap_mean76 = Cap_sum76 / Cap_cnt76
                if df3_freq_Cap77 > df3_Freq_Cap:
                    Cap_sum77 += df3_freq_Cap77
                    Cap_cnt77 += 1
                    Cap_mean77 = Cap_sum77 / Cap_cnt77
                if df3_freq_Cap78 > df3_Freq_Cap:
                    Cap_sum78 += df3_freq_Cap78
                    Cap_cnt78 += 1
                    Cap_mean78 = Cap_sum78 / Cap_cnt78
                if df3_freq_Cap79 > df3_Freq_Cap:
                    Cap_sum79 += df3_freq_Cap79
                    Cap_cnt79 += 1
                    Cap_mean79 = Cap_sum79 / Cap_cnt79
                if df3_freq_Cap80 > df3_Freq_Cap:
                    Cap_sum80 += df3_freq_Cap80
                    Cap_cnt80 += 1
                    Cap_mean80 = Cap_sum80 / Cap_cnt80
                if df3_freq_Cap81 > df3_Freq_Cap:
                    Cap_sum81 += df3_freq_Cap81
                    Cap_cnt81 += 1
                    Cap_mean81 = Cap_sum81 / Cap_cnt81
                if df3_freq_Cap82 > df3_Freq_Cap:
                    Cap_sum82 += df3_freq_Cap82
                    Cap_cnt82 += 1
                    Cap_mean82 = Cap_sum82 / Cap_cnt82
                if df3_freq_Cap83 > df3_Freq_Cap:
                    Cap_sum83 += df3_freq_Cap83
                    Cap_cnt83 += 1
                    Cap_mean83 = Cap_sum83 / Cap_cnt83
                if df3_freq_Cap84 > df3_Freq_Cap:
                    Cap_sum84 += df3_freq_Cap84
                    Cap_cnt84 += 1
                    Cap_mean84 = Cap_sum84 / Cap_cnt84
                if df3_freq_Cap85 > df3_Freq_Cap:
                    Cap_sum85 += df3_freq_Cap85
                    Cap_cnt85 += 1
                    Cap_mean85 = Cap_sum85 / Cap_cnt85
                if df3_freq_Cap86 > df3_Freq_Cap:
                    Cap_sum86 += df3_freq_Cap86
                    Cap_cnt86 += 1
                    Cap_mean86 = Cap_sum86 / Cap_cnt86
                if df3_freq_Cap87 > df3_Freq_Cap:
                    Cap_sum87 += df3_freq_Cap87
                    Cap_cnt87 += 1
                    Cap_mean87 = Cap_sum87 / Cap_cnt87
                if df3_freq_Cap88 > df3_Freq_Cap:
                    Cap_sum88 += df3_freq_Cap88
                    Cap_cnt88 += 1
                    Cap_mean88 = Cap_sum88 / Cap_cnt88
                if df3_freq_Cap89 > df3_Freq_Cap:
                    Cap_sum89 += df3_freq_Cap89
                    Cap_cnt89 += 1
                    Cap_mean89 = Cap_sum89 / Cap_cnt89
                if df3_freq_Cap90 > df3_Freq_Cap:
                    Cap_sum90 += df3_freq_Cap90
                    Cap_cnt90 += 1
                    Cap_mean90 = Cap_sum90 / Cap_cnt90
                if df3_freq_Cap91 > df3_Freq_Cap:
                    Cap_sum91 += df3_freq_Cap91
                    Cap_cnt91 += 1
                    Cap_mean91 = Cap_sum91 / Cap_cnt91
                if df3_freq_Cap92 > df3_Freq_Cap:
                    Cap_sum92 += df3_freq_Cap92
                    Cap_cnt92 += 1
                    Cap_mean92 = Cap_sum92 / Cap_cnt92
                if df3_freq_Cap93 > df3_Freq_Cap:
                    Cap_sum93 += df3_freq_Cap93
                    Cap_cnt93 += 1
                    Cap_mean93 = Cap_sum93 / Cap_cnt93
                if df3_freq_Cap94 > df3_Freq_Cap:
                    Cap_sum94 += df3_freq_Cap94
                    Cap_cnt94 += 1
                    Cap_mean94 = Cap_sum94 / Cap_cnt94
                if df3_freq_Cap95 > df3_Freq_Cap:
                    Cap_sum95 += df3_freq_Cap95
                    Cap_cnt95 += 1
                    Cap_mean95 = Cap_sum95 / Cap_cnt95
                if df3_freq_Cap96 > df3_Freq_Cap:
                    Cap_sum96 += df3_freq_Cap96
                    Cap_cnt96 += 1
                    Cap_mean96 = Cap_sum96 / Cap_cnt96
                if df3_freq_Cap97 > df3_Freq_Cap:
                    Cap_sum97 += df3_freq_Cap97
                    Cap_cnt97 += 1
                    Cap_mean97 = Cap_sum97 / Cap_cnt97
                if df3_freq_Cap98 > df3_Freq_Cap:
                    Cap_sum98 += df3_freq_Cap98
                    Cap_cnt98 += 1
                    Cap_mean98 = Cap_sum98 / Cap_cnt98
                if df3_freq_Cap99 > df3_Freq_Cap:
                    Cap_sum99 += df3_freq_Cap99
                    Cap_cnt99 += 1
                    Cap_mean99 = Cap_sum99 / Cap_cnt99
                if df3_freq_Cap100 > df3_Freq_Cap:
                    Cap_sum100 += df3_freq_Cap100
                    Cap_cnt100 += 1
                    Cap_mean100 = Cap_sum100 / Cap_cnt100
                else:
                    pass
                    
                df3_Cap_mean = [Cap_mean1,Cap_mean2,Cap_mean3,Cap_mean4,Cap_mean5,Cap_mean6,Cap_mean7,Cap_mean8,Cap_mean9,Cap_mean10,Cap_mean11,\
                     Cap_mean12,Cap_mean13,Cap_mean14,Cap_mean15,Cap_mean16,Cap_mean17,Cap_mean18,Cap_mean19,Cap_mean20,Cap_mean21,Cap_mean22,\
                     Cap_mean23,Cap_mean24,Cap_mean25,Cap_mean26,Cap_mean27,Cap_mean28,Cap_mean29,Cap_mean30,Cap_mean31,Cap_mean32,Cap_mean33,\
                     Cap_mean34,Cap_mean35,Cap_mean36,Cap_mean37,Cap_mean38,Cap_mean39,Cap_mean40,Cap_mean41,Cap_mean42,Cap_mean43,Cap_mean44,\
                     Cap_mean45,Cap_mean46,Cap_mean47,Cap_mean48,Cap_mean49,Cap_mean50,Cap_mean51,Cap_mean52,Cap_mean53,Cap_mean54,Cap_mean55,\
                     Cap_mean56,Cap_mean57,Cap_mean58,Cap_mean59,Cap_mean60,Cap_mean61,Cap_mean62,Cap_mean63,Cap_mean64,Cap_mean65,Cap_mean66,\
                     Cap_mean67,Cap_mean68,Cap_mean69,Cap_mean70,Cap_mean71,Cap_mean72,Cap_mean73,Cap_mean74,Cap_mean75,Cap_mean76,Cap_mean77,\
                     Cap_mean78,Cap_mean79,Cap_mean80,Cap_mean81,Cap_mean82,Cap_mean83,Cap_mean84,Cap_mean85,Cap_mean86,Cap_mean87,Cap_mean88,\
                     Cap_mean89,Cap_mean90,Cap_mean91,Cap_mean92,Cap_mean93,Cap_mean94,Cap_mean95,Cap_mean96,Cap_mean97,Cap_mean98,Cap_mean99,\
                     Cap_mean100]
                     
        df3_Cap_mean = [round(i,1) for i in df3_Cap_mean]#list取1位小数
        
        
    '''获取含坐标的频率平均值'''
    def Get_Freq_Mean_csv(self):
        global df3_XY, df3_Cap_XY, df3_Cap_mean
        df3_CapXY = df3_XY
        Cap_mean = df3_Cap_mean    #df3_Cap_mean含频率值等于0的通道，df3_Cap_mean1,2不含
        df3_CapXY.insert(3,'Cap_mean',Cap_mean)
        df3_Cap_XY = df3_CapXY[df3_CapXY.columns[0:4]] #获取前4列['Bin','X','Y','Cap_mean'],根据需求修改
    
            
    '''获取非负坐标频率测试项Map图'''
    def Get_Freq_Map(self):
        global df3_cnt, df3_Cap_XY
        
        #颜色定义
        styleLightGreenBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;')
        styleYellowBkg = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;')
        styleLightBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;')
        styleIceBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue;')
        styleSkyBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;')
        styleOrangeBkg = xlwt.easyxf('pattern: pattern solid, fore_colour orange;')
        stylePaleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;')
        
        if df3_cnt == 46:
            arr = list([4 for x in range(46)])  #产生值为4的list列表
            Cap_X = []
            Cap_Y = []
        
            Cap_X = df3_Cap_XY['X'].tolist()
            Cap_Y = df3_Cap_XY['Y'].tolist()
            Cap_X = list(map(int,Cap_X))
            Cap_Y = list(map(int,Cap_Y))
            Cap_Xlist = []
            Cap_Ylist = []
            Cap_Xarrlist = []
            Cap_Yarrlist = []
            for i in range(46):
                Cap_Xlist = Cap_X[i] + arr[i]
                Cap_Xlist = str(Cap_Xlist)  
                Cap_Xarrlist.append(Cap_Xlist)
                
            for j in range(46):
                Cap_Ylist = Cap_Y[j] + arr[j]
                Cap_Ylist = str(Cap_Ylist)
                Cap_Yarrlist.append(Cap_Ylist)
                
        elif df3_cnt == 100:
            arr = list([8 for x in range(100)])  #产生值为8的list列表
            Cap_X = []
            Cap_Y = []
        
            Cap_X = df3_Cap_XY['X'].tolist()
            Cap_Y = df3_Cap_XY['Y'].tolist()
            Cap_X = list(map(int,Cap_X))
            Cap_Y = list(map(int,Cap_Y))
            Cap_Xlist = []
            Cap_Ylist = []
            Cap_Xarrlist = []
            Cap_Yarrlist = []
            for i in range(100):
                Cap_Xlist = Cap_X[i] + arr[i]
                Cap_Xlist = str(Cap_Xlist)  
                Cap_Xarrlist.append(Cap_Xlist)
                
            for j in range(100):
                Cap_Ylist = Cap_Y[j] + arr[j]
                Cap_Ylist = str(Cap_Ylist)
                Cap_Yarrlist.append(Cap_Ylist)
        else:
            pass
            
        df3_Cap_XY.insert(3,'Cap_X',Cap_Xarrlist)
        df3_Cap_XY.insert(4,'Cap_Y',Cap_Yarrlist)
        
        df3_Cap_XY = df3_Cap_XY[df3_Cap_XY.columns[3:6]]
    
        #删除上次生成的文件,避免出现无法修改数据错误
        RCH_PATH = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\Freq-Mapdata.txt'
        RCH_PATH1 = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\Freq-Mapdata.xls'
        if os.path.exists(RCH_PATH):
            os.remove(RCH_PATH)
        elif os.path.exists(RCH_PATH1):
            os.remove(RCH_PATH1)
        else:
            pass
            
        #存储为txt文件
        with open(RCH_PATH,'a+',encoding='utf-8') as f1:
            for line in df3_Cap_XY.values:
                f1.write((str(line[0])+'\t'+str(line[1])+'\t'+str(line[2])+'\n'))
        workbook1 = xlwt.Workbook()
        worksheet1 = workbook1.add_sheet('Freq-Map')
        
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
#        print("频率Freq测试完成!")
    
    @pyqtSlot()
    def on_btn_Freq_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
#        print("点击频率测试项按钮")
        self.Get_csv()
        self.Get_New_csv()
        self.Wafer_Dies()
        self.Get_Freq_raw_data()
        self.Get_Freq_mean()
        
        self.Get_Freq_Mean_csv()
        self.Get_Freq_Map()
        print("频率测试项完成!")
        
        
        
