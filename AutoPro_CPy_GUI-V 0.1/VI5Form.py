# -*- coding: utf-8 -*-

"""
Module implementing VI5Form.
"""

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget

from Ui_VI5Form import Ui_Form
import pandas as pd
import os
import xlwt


class VI5Form(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    df1 = []#全局变量定义
    df3_new = []
    df3 = []
    df3_XY = []
    df3_cnt = []
    df3_VCH = []
    df3_VI5_VCH = []
    
    df3_VI5_mean_value = []
    df3_VI5_XY =[]
    
    
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget (defaults to None)
        @type QWidget (optional)
        """
        super(VI5Form, self).__init__(parent)
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
            
    '''获取V/I特性原始数据'''
    def Get_VI_data(self):
        global df3, df3_new, df3_VCH
        VCH_cols_num = [i for i,x in enumerate(df3.columns) if x.find('VCH')!=-1]
        VCH_cols = df3.columns[VCH_cols_num]
        for VCHcols_num in VCH_cols:
            df3_new[VCHcols_num] = df3_new[VCHcols_num].astype('float64')#类型转换
            
        df3_VCH = df3_new[VCH_cols]
        
    '''获取5mA电流条件下，V/I特性原始数据'''
    def Get_VI5_data(self):
        global df3_VCH, df3_VI5_VCH
        VI5_cols_num = [i for i,x in enumerate(df3_VCH.columns) if x.find('_1')!=-1]
        VI5_cols = df3_VCH.columns[VI5_cols_num]
        df3_VI5_VCH = df3_VCH[VI5_cols]
        df3_VI5_VCH = df3_VI5_VCH.drop(['VCH1_10','VCH2_10','VCH3_10','VCH4_10','VCH5_10',\
                                'VCH6_10','VCH7_10','VCH8_10','VCH9_10',\
                                'VCH10_10','VCH11_10','VCH12_10','VCH13_10',\
                                'VCH14_10','VCH15_10','VCH16_10','VCH17_10',\
                                'VCH18_10','VCH19_10','VCH20_10','VCH21_10',\
                                'VCH22_10','VCH23_10','VCH24_10','VCH25_10',\
                                'VCH26_10','VCH27_10','VCH28_10','VCH29_10',\
                                'VCH30_10','VCH31_10','VCH32_10','VCH33_10',\
                                'VCH34_10','VCH35_10','VCH36_10','VCH37_10',\
                                'VCH38_10','VCH39_10','VCH40_10','VCH41_10',\
                                'VCH42_10','VCH43_10','VCH44_10','VCH45_10',\
                                'VCH46_10','VCH47_10','VCH48_10','VCH49_10',\
                                'VCH50_10','VCH51_10','VCH52_10'],axis = 1)
                                
    '''获取5mA电流条件下，同颗芯片52通道电压平均值'''
    def Get_VI5_mean(self):
        global df3_VI5_VCH, df3_VI5_mean_value
        VI5_index = df3_VI5_VCH.index-18
        VI5_cols = df3_VI5_VCH.columns
        VI5_sum = 0.0
        VI5_mean = 0.0
        VI5_cnt = 0
        df3_VI5_mean = []
        
        for VI5_num in VI5_index:
            for VI5_num1 in VI5_cols:
                VI5_value = df3_VI5_VCH.iloc[VI5_num][VI5_num1]
                if VI5_value < 18.000:
                    VI5_sum += VI5_value
                    VI5_cnt += 1
                    VI5_mean = VI5_sum / VI5_cnt
                else:
                    pass
                    
            df3_VI5_mean.append(VI5_mean)
            df3_VI5_mean = [round(i,1) for i in df3_VI5_mean]
            VI5_sum = 0.0
            VI5_mean = 0.0
            VI5_cnt = 0
            
        df3_VI5_mean_value = df3_VI5_mean
            
    '''获取5mA电流条件下，电压平均值表格csv'''
    def Get_VI5_Mean_csv(self):
        global df3_XY,  df3_VI5_XY, df3_VI5_mean_value
        df3_VI5XY = df3_XY
        VI5_mean = df3_VI5_mean_value
        df3_VI5XY.insert(3,'VI5_mean',VI5_mean)
        df3_VI5_XY = df3_VI5XY[df3_VI5XY.columns[0:4]] #获取前4列['Bin','X','Y','Cap_mean'],根据需求修改
        
    '''获取5mA电流条件下，非负坐标电压平均值Map'''
    def Get_VI5_Map(self):
        global df3_cnt, df3_VI5_XY
        
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
            VI5_X = []
            VI5_Y = []
        
            VI5_X = df3_VI5_XY['X'].tolist()
            VI5_Y = df3_VI5_XY['Y'].tolist()
            VI5_X = list(map(int,VI5_X))
            VI5_Y = list(map(int,VI5_Y))
            VI5_Xlist = []
            VI5_Ylist = []
            VI5_Xarrlist = []
            VI5_Yarrlist = []
            for i in range(46):
                VI5_Xlist = VI5_X[i] + arr1[i]
                VI5_Xlist = str(VI5_Xlist)  
                VI5_Xarrlist.append(VI5_Xlist)
                
            for j in range(46):
                VI5_Ylist = VI5_Y[j] + arr1[j]
                VI5_Ylist = str(VI5_Ylist)
                VI5_Yarrlist.append(VI5_Ylist)
                
        elif df3_cnt == 100:
            arr1 = list([8 for x in range(100)])  #产生值为4的list列表
            VI5_X = []
            VI5_Y = []
        
            VI5_X = df3_VI5_XY['X'].tolist()
            VI5_Y = df3_VI5_XY['Y'].tolist()
            VI5_X = list(map(int,VI5_X))
            VI5_Y = list(map(int,VI5_Y))
            VI5_Xlist = []
            VI5_Ylist = []
            VI5_Xarrlist = []
            VI5_Yarrlist = []
            for i in range(100):
                VI5_Xlist = VI5_X[i] + arr1[i]
                VI5_Xlist = str(VI5_Xlist)  
                VI5_Xarrlist.append(VI5_Xlist)
                
            for j in range(100):
                VI5_Ylist = VI5_Y[j] + arr1[j]
                VI5_Ylist = str(VI5_Ylist)
                VI5_Yarrlist.append(VI5_Ylist)
        else:
            pass
            
        df3_VI5_XY.insert(3,'VI5_X',VI5_Xarrlist)
        df3_VI5_XY.insert(4,'VI5_Y',VI5_Yarrlist)
        
        df3_VI5_XY = df3_VI5_XY[df3_VI5_XY.columns[3:6]]
        
        #删除上次生成的文件,避免出现无法修改数据错误
        RCH_PATH = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\VI5-Mapdata.txt'
        RCH_PATH1 = os.path.abspath(os.path.dirname(__file__))+'\csv_data\\VI5-Mapdata.xls'
        if os.path.exists(RCH_PATH):
            os.remove(RCH_PATH)
        elif os.path.exists(RCH_PATH1):
            os.remove(RCH_PATH1)
        else:
            pass
        
        #存储为txt文件
        with open(RCH_PATH,'a+',encoding='utf-8') as f1:
            for line in df3_VI5_XY.values:
                f1.write((str(line[0])+'\t'+str(line[1])+'\t'+str(line[2])+'\n'))
        workbook1 = xlwt.Workbook()
        worksheet1 = workbook1.add_sheet('VI5-Map')
        
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
#        print("伏安特性曲线VI5完成!")
    
    
    @pyqtSlot()
    def on_btn_VI5_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
#        print("点击伏安特性曲线测试按钮")
        self.Get_csv()
        self.Get_New_csv()
        self.Wafer_Dies()
        self.Get_VI_data()
        self.Get_VI5_data()
        self.Get_VI5_mean()
        
        self.Get_VI5_Mean_csv()
        self.Get_VI5_Map()
        print("伏安特性测试项完成!")
