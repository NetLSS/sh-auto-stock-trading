import sys
import os
from PyQt5.QtWidgets import *
from PyQt5 import uic

from PyQt5.QtGui import *
from PyQt5.QAxContainer import *

from pandas import Series, DataFrame
import matplotlib.pyplot as plt
import pandas as pd
import mpl_finance as matfin
import matplotlib.ticker as ticker

import numpy as np

import win32com.shell.shell as shell


# 관리자 권한 획득
if True:
    ASADMIN = 'asadmin'

    if sys.argv[-1] != ASADMIN:
        script = os.path.abspath(sys.argv[0])
        params = ' '.join([script] + sys.argv[1:] + [ASADMIN])
        shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
        sys.exit(0)

# ui 루트 경로 지정
os.chdir(r"C:\1.MyPersonal\Project\auto-stock-trading\source\simple_price_checker")


#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("sslee.ui")[0]

# 종목
stock_item_1 = "017180" # 명문 제약
stock_item_2 = "058820" # CMG 제약
 
#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.RequestRealTimeStock_click_flag = False
        self.ReceiveRealData_click_flag = False
        
        # #버튼에 기능을 연결하는 코드
        # self.btn_1.clicked.connect(self.button1Function)
        # self.btn_2.clicked.connect(self.button2Function)

        # stock
         # 일반,실시간 TR OCX
        self.CommTR = QAxWidget("GIEXPERTCONTROL.GiExpertControlCtrl.1")
        self.CommReal = QAxWidget("GIEXPERTCONTROL.GiExpertControlCtrl.1")

        self.CommTR.ReceiveData.connect(self.ReceiveTRData)
        self.CommReal.ReceiveRTData.connect(self.ReceiveRealData)

        self.pushButton_Re.clicked.connect(self.pushButton_Re_clicked)
         # TR ID를 저장해놓고 처리할 딕셔너리 생성
        self.rqid = {}
    
    

    def ReceiveTRData(self, nID):
        TRName = self.rqid.get(nID)
        if TRName == "SC":
            self.TR_SC_process()
        elif TRName == "CANDLE":
            # self.DrawCandleChart()
            pass
        elif TRName == "LINE":
            # self.DrawChart()
            pass
        elif TRName == "stock_mst":
            # codelist = []
            # count = self.CommTR.dynamicCall("GetMultiRowCount()")
            # for i in range(0, count):
            #     gubun = self.CommTR.dynamicCall("GetMultiData(int, QString)", i, 2)
            #     if gubun == "0":
            #         code = self.CommTR.dynamicCall("GetMultiData(int, QString)", i, 1)
            #         name = self.CommTR.dynamicCall("GetMultiData(int, QString)", i, 3)
            #         codelist.append(code + " : " + name)
            # self.listWidget.addItems(codelist)
            pass
        elif TRName == "SB":
            # print("SB")
            # codelist = []
            # count = self.CommTR.dynamicCall("GetMultiRowCount()")
            # 종목 = self.CommTR.dynamicCall("GetSingleData(int)", 5)
            # print(f"SB: {count} {종목}")
            pass
        self.rqid.__delitem__(nID)

    def ReceiveRealData(self, RealType):
        
        self.ReceiveRealData_click_flag = not slef.ReceiveRealData_click_flag
        if self.ReceiveRealData_click_flag:
            self.pushButton_Re.setText("RE!")
        else:
            self.pushButton_Re.setText("RE")


        if RealType == "SC":
            DATA = {}
            DATA['ISIN_CODE'] = self.CommReal.dynamicCall("GetSingleData(int)", 0)  # 표준코드
            DATA['CODE'] = self.CommReal.dynamicCall("GetSingleData(int)", 1)  # 단축코드
            DATA['Time'] = self.CommReal.dynamicCall("GetSingleData(int)", 2)  # 채결시간
            DATA['Close'] = self.CommReal.dynamicCall("GetSingleData(int)", 3)  # 현재가
            DATA['a'] = self.CommReal.dynamicCall("GetSingleData(int)", 4)  # 전일대비 구분
            DATA['b'] = self.CommReal.dynamicCall("GetSingleData(int)", 5)  # 전일대비
            DATA['Increase'] = str(self.CommReal.dynamicCall("GetSingleData(int)", 6))  # 전일대비율%
            DATA['Vol'] = self.CommReal.dynamicCall("GetSingleData(int)", 7)  # 누적거래량
            DATA['TRADING_VALUE'] = self.CommReal.dynamicCall("GetSingleData(int)", 8)  # 누적거래대금
            DATA['ContQty'] = self.CommReal.dynamicCall("GetSingleData(int)", 9)  # 단위채결량
            DATA['Open'] = self.CommReal.dynamicCall("GetSingleData(int)", 10)  # 시가
            DATA['High'] = self.CommReal.dynamicCall("GetSingleData(int)", 11)  # 고가
            DATA['Low'] = self.CommReal.dynamicCall("GetSingleData(int)", 12)  # 저가
        
        self.UpdateUI(DATA)
        # for i in range (1, 10):
        #     x = self.CommReal.dynamicCall("GetSingleData(int)", i)
        #     self.tableWidget.setItem(i, 0, QTableWidgetItem(x))

    def UpdateUI(self, DATA):
        if str(DATA['CODE']) == stock_item_1:
            price = DATA.get("Close")
            if price is not None:
                self.lineEdit_price1.setText(price)

            increase = DATA.get("Increase")
            if increase is not None:
                self.lineEdit_increase1.setText(increase)

        elif str(DATA['CODE']) == stock_item_2:
            price = DATA.get("Close")
            if price is not None:
                self.lineEdit_price2.setText(price)

            increase = DATA.get("Increase")
            if increase is not None:
                self.lineEdit_increase2.setText(increase)
        
    def TR_SC_process(self):
        """
        TR 에서 오는 SC 처리
        """
        DATA = {}
        DATA['ISIN_CODE'] = self.CommTR.dynamicCall("GetSingleData(int)", 0) # 표준코드
        DATA['CODE'] = self.CommTR.dynamicCall("GetSingleData(int)", 1)       # 단축코드
        DATA['Time'] = self.CommTR.dynamicCall("GetSingleData(int)", 2)       # 채결시간
        DATA['Close'] = self.CommTR.dynamicCall("GetSingleData(int)", 3)      # 현재가
        DATA['Vol'] = self.CommTR.dynamicCall("GetSingleData(int)", 7)        # 누적거래량
        DATA['TRADING_VALUE'] = self.CommTR.dynamicCall("GetSingleData(int)", 8)  # 누적거래대금
        DATA['ContQty'] = self.CommTR.dynamicCall("GetSingleData(int)", 9)   # 단위채결량
        DATA['Open'] = self.CommTR.dynamicCall("GetSingleData(int)", 10)      # 시가
        DATA['High'] = self.CommTR.dynamicCall("GetSingleData(int)", 11)      # 고가
        DATA['Low'] = self.CommTR.dynamicCall("GetSingleData(int)", 12)       # 저가
        print(DATA)
        self.UpdateUI(DATA)

    def RequestStockList(self):
        self.CommTR.dynamicCall("SetQueryName(QString)", "stock_mst")
        nResult = self.CommTR.dynamicCall("RequestData()")
        self.rqid[nResult] = "stock_mst"

        self.CommTR.dynamicCall("SetQueryName(QString)", "SB")
        self.CommTR.dynamicCall("SetSingleData(int, QString)", 0, "055550") # 005933:삼성전자
        nResult = self.CommTR.dynamicCall("RequestData()")
        self.rqid[nResult] = "SB"

        self.CommTR.dynamicCall("SetQueryName(QString)", "SC")
        self.CommTR.dynamicCall("SetSingleData(int, QString)", 0, "055550") # 005933:삼성전자
        nResult = self.CommTR.dynamicCall("RequestData()")
        self.rqid[nResult] = "SC"

    def RequestTRStock(self):
        self.CommReal.dynamicCall("UnRequestRTRegAll()")
        
        ret = self.CommTR.dynamicCall("SetQueryName(QString)", "SC")
        ret = self.CommTR.dynamicCall("SetSingleData(int, QString)", 0, stock_item_1) # 인풋 : 단축코드
        rqid = self.CommTR.dynamicCall("RequestData()")
        self.rqid[rqid] = "SC"

        ret = self.CommTR.dynamicCall("SetQueryName(QString)", "SC")
        ret = self.CommTR.dynamicCall("SetSingleData(int, QString)", 0, stock_item_2) # 인풋 : 단축코드
        rqid = self.CommTR.dynamicCall("RequestData()")
        self.rqid[rqid] = "SC"
    

    def RequestRealTimeStock(self):
        self.RequestRealTimeStock_click_flag = not self.RequestRealTimeStock_click_flag
        if self.RequestRealTimeStock_click_flag:
            self.pushButton_Re.setText("RE.")
        else:
            self.pushButton_Re.setText("RE")

        ret = self.CommReal.dynamicCall("RequestRTReg(QString, QString)", "SC", stock_item_1)
        ret = self.CommReal.dynamicCall("RequestRTReg(QString, QString)", "SC", stock_item_2)

    def pushButton_Re_clicked(self):
        self.RequestTRStock()
        self.RequestRealTimeStock()


if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass() 
    myWindow.RequestTRStock()
    myWindow.RequestRealTimeStock()
    myWindow.show()
    app.exec_()