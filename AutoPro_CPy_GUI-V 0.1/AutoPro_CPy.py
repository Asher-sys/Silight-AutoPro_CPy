# -*- coding: utf-8 -*-

"""
Module implementing AutoProMainWindow.
"""
from PyQt5 import QtWidgets
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QMessageBox
from PyQt5.QtGui import QStatusTipEvent

from Ui_AutoPro_CPy import Ui_MainWindow
from SlotResForm import SlotResForm
from SlotSlopeForm import SlotSlopeForm
from PIForm import PIForm
from LeakageIForm import LeakageIForm
from FreqForm import FreqForm
from VI5Form import VI5Form


class AutoProMainWindow(QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget (defaults to None)
        @type QWidget (optional)
        """
        super(AutoProMainWindow, self).__init__(parent)
        self.setupUi(self)
        
        self.statusBar.showMessage("版权：上海信及光子集成技术有限公司")
        
    def event(self, QEvent):
        if QEvent.type() == QEvent.StatusTip:
            if QEvent.tip() == "":
                QEvent = QStatusTipEvent("版权：上海信及光子集成技术有限公司")
        return super().event(QEvent)
    
    @pyqtSlot()
    def on_actionSlotSlope_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        slotslopeform.show()
    
    @pyqtSlot()
    def on_actionVI5_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        vi5form.show()
    
    @pyqtSlot()
    def on_actionSlotRes_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        slotresform.show()
    
    @pyqtSlot()
    def on_actionPI_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        piform.show()
    
    @pyqtSlot()
    def on_actionLeakageI_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        leakageiform.show()
    
    @pyqtSlot()
    def on_actionFreq_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        freqform.show()
    
    @pyqtSlot()
    def on_actionExit_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        print("Click Menu Exit")
        reply = QMessageBox.question(self,"提示","确定要退出系统吗?", QMessageBox.Yes|QMessageBox.No)
        if reply == QMessageBox.Yes:
            sys.exit(0)
    
    @pyqtSlot()
    def on_actionAbout_triggered(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #raise NotImplementedError
        reply = QMessageBox.question(self,"关于","AutoPro_CPy 软件版本 V0.1", QMessageBox.Yes|QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.close()
        else:
            self.close()
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    autopro = AutoProMainWindow()
    autopro.show()
    
    #实例化子窗体类
    slotresform = SlotResForm()
    slotslopeform = SlotSlopeForm()
    piform = PIForm()
    leakageiform = LeakageIForm()
    freqform = FreqForm()
    vi5form = VI5Form()
    sys.exit(app.exec_())
