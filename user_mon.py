
# coding: utf-8

# ■ 사용된 PLUS OBJECT:
# 
#   - CpSysDib.CssStgList : 전략 리스트 조회 (예제 또는 사용자 전략 선택 가능)
# 
#   - CpSysDib.CssStgFind : 특정 전략조건에 해당하는 종목 리스트 조회 

# In[1]:


import sys
from PyQt5.QtWidgets import *
import win32com.client
import pandas as pd
import os
from pandas import DataFrame
from time import sleep

import ChatBotModel
from tabulate import tabulate


# In[2]:


# 텔레그램 봇 생성하기
BUSbot = ChatBotModel.Bot2ndBUS()

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

    def OnReceived(self):
        pbData = {}
        # 실시간 종목검색 감시 처리 
        if self.name == 'cssalert':
            pbData['전략ID'] = self.client.GetHeaderValue(0)
            pbData['감시일련번호'] = self.client.GetHeaderValue(1)
            code = pbData['code'] = self.client.GetHeaderValue(2)
            pbData['종목명'] = name = g_objCodeMgr.CodeToName(code)

            inoutflag = self.client.GetHeaderValue(3)
            if (ord('1') == inoutflag):
                pbData['INOUT'] = '진입'
            elif (ord('2') == inoutflag):
                pbData['INOUT'] = '퇴출'
            pbData['시각'] = self.client.GetHeaderValue(4)
            pbData['현재가'] = self.client.GetHeaderValue(5)
            self.caller.checkRealtimeStg(pbData)


class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def Subscribe(self, var, caller):
        if self.bIsSB:
            self.Unsubscribe()

        if (len(var) > 0):
            self.obj.SetInputValue(0, var)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, caller)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False


# CpPBCssAlert: 종목검색 실시간 PB 클래스 
class CpPBCssAlert(CpPublish):
    def __init__(self):
        super().__init__('cssalert', 'CpSysDib.CssAlert')


# Cp8537 : 종목검색 전략 조회
class Cp8537:
    def __init__(self):
        self.objpb = CpPBCssAlert()
        self.bisSB = False
        self.monList = {}

    def __del__(self):
        self.Clear()

    def Clear(self):
        self.stopAllStgControl()
        if self.bisSB:
            self.objpb.Unsubscribe()
            self.bisSB = False

    def requestList(self, sel, caller):
        objRq = win32com.client.Dispatch("CpSysDib.CssStgList")  #전략조회 객체

        # 예제 전략에서 전략 리스트를 가져옵니다.
        if (sel == '예제'):
            objRq.SetInputValue(0, ord('0'))  # '0' : 예제전략, '1': 나의전략
        else:
            objRq.SetInputValue(0, ord('1'))  # '0' : 예제전략, '1': 나의전략
        
        objRq.BlockRequest()

        # 통신 및 통신 에러 처리
        rqStatus = objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            return False

        cnt = objRq.GetHeaderValue(0)  # 0 - (long) 전략 목록 수
        flag = objRq.GetHeaderValue(1)  # 1 - (char) 요청구분

        for i in range(cnt):
            item = {}
            item['전략명'] = objRq.GetDataValue(0, i)
            item['ID'] = objRq.GetDataValue(1, i)
            item['전략등록일시'] = objRq.GetDataValue(2, i)
            item['작성자필명'] = objRq.GetDataValue(3, i)
            item['평균종목수'] = objRq.GetDataValue(4, i)
            item['평균승률'] = objRq.GetDataValue(5, i)
            item['평균수익'] = objRq.GetDataValue(6, i)
            caller.StgList[item['전략명']] = item

        return True

    def requestStgID(self, id, caller):
        objRq = None
        objRq = win32com.client.Dispatch("CpSysDib.CssStgFind")
        objRq.SetInputValue(0, id)  # 전략 id 요청
        objRq.BlockRequest()
        # 통신 및 통신 에러 처리
        rqStatus = objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            return False

        cnt = objRq.GetHeaderValue(0)  # 0 - (long) 검색된 결과 종목 수
        totcnt = objRq.GetHeaderValue(1)  # 1 - (long) 총 검색 종목 수
        stime = objRq.GetHeaderValue(2)  # 2 - (string) 검색시간
        print('검색된 종목수:', cnt, '전체종목수:', totcnt, '검색시간:', stime)
        
        caller.dataStgList = []
        
        for i in range(cnt):
            # caller.dataStgList.append = (index, code, 종목명)
            code = objRq.GetDataValue(0, i)
            name = g_objCodeMgr.CodeToName(code)
            caller.dataStgList.append([i, code, name])

        return True

    def requestMonitorID(self, id, caller):
        objRq = win32com.client.Dispatch("CpSysDib.CssWatchStgSubscribe")
        objRq.SetInputValue(0, id)  # 전략 id 요청
        objRq.BlockRequest()

        # 통신 및 통신 에러 처리
        rqStatus = objRq.GetDibStatus()

        if rqStatus != 0:
            rqRet = objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            return False

        caller.monID = objRq.GetHeaderValue(0)

        if caller.monID== 0:
            print('감시 일련번호 구하기 실패')
            return False

        # monID - 전략 감시를 위한 일련번호를 구해온다.
        # 현재 감시되는 전략이 없다면 감시일련번호로 1을 리턴하고,
        # 현재 감시되는 전략이 있다면 각 통신 ID에 대응되는 새로운 일련번호를 리턴한다.
        return True

    def requestStgControl(self, id, monID, bStart):
        objRq = win32com.client.Dispatch("CpSysDib.CssWatchStgControl")
        objRq.SetInputValue(0, id)  # 전략 id 요청
        objRq.SetInputValue(1, monID)  # 감시일련번호
        
        # 전략 감시시 너무 많은 진입/탈출이 반복적으로 발생되어 전략감시는 사용하지 않음
        bStart = False  # 임의로 실시간 전략감시 취소
        
        # 전략 감시
        if bStart == True:
            objRq.SetInputValue(2, ord('1'))  # 감시시작
            print('전략감시 시작 요청 ', '전략 ID:', id, '감시일련번호', monID)
        else:
            objRq.SetInputValue(2, ord('3'))  # 감시취소
            print('전략감시 취소 요청 ', '전략 ID:', id, '감시일련번호', monID)
        objRq.BlockRequest()

        # 통신 및 통신 에러 처리
        rqStatus = objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            return (False, '')

        status = objRq.GetHeaderValue(0)

        if status == 0:
            print('전략감시상태: 초기상태')
        elif status == 1:
            print('전략감시상태: 감시중')
        elif status == 2:
            print('전략감시상태: 감시중단')
        elif status == 3:
            print('전략감시상태: 등록취소')

        # event 수신 요청 - 요청 중이 아닌 경우에만 요청
        if self.bisSB == False:
            self.objpb.Subscribe('', self)
            self.bisSB = True

        # 진행 중인 전략들 저장
        if bStart == True:
            self.monList[id] = monID
        else:
            if id in self.monList:
                del self.monList[id]
        print (self.monList) ##

        return (True, status)

    def stopAllStgControl(self):
        delitem = []
        for id, monID in self.monList.items():
            delitem.append((id, monID))

        for item in delitem:
            self.requestStgControl(item[0], item[1], False)

        print(len(self.monList))

    def checkRealtimeStg(self, pbData):
        # 감시중인 전략인 경우만 체크
        id = pbData['전략ID']
        monID = pbData['감시일련번호']
        if not (id in self.monList):
            return

        if (monID != self.monList[id]):
            return

        #텔레그램 봇으로 메세지보내기
        df = DataFrame([pbData])  
        message = tabulate(df, headers='keys', tablefmt='simple')
        message = "<pre>{}</pre>".format(message)
        BUSbot.sendMessage2html(message)
        
        # 전략ID, 감시일련번호, code, 종목명, INOUT, 시가, 현재가
        print ("텔레그램 봇으로 보냄")
        print(pbData)


# In[3]:


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("2ndBUS")
        
        #텔레그램봇 메세지 보내기
        try:
            BUSbot.sendMessage("유저전략 검색을 시작합니다.")
        except Exception as ex:
            print ("텔레그램 오류가 발생했습니다.", ex)
            pass  # telegram 처음 첫 시작시 timeout error가 발생하는 경우있어, 발생시 무시고하고 진행함

        # 전략명 지정하여 검색된 종목
        self.stgName = ["볼린저밴드 I","볼린저밴드 II.8","볼린저밴드 II.2","볼린저밴드 III.w2"]
        self.obj8537 = Cp8537()
        self.dataStgList = []
        self.StgList = {}
        self.monID = 0
        self.id = []

        self.listMyStrategy()  # 전략리스트 조회
        
        ret = self.total_check()
        if ret == False:
            self.obj8537.Clear()  # 기존 감시 종료
            self.listMyStrategy()  # 전략리스트 조회
        
        self.monitor_stg()  # 실시간 전략감시
        
        # 윈도우 스케줄러에서 반복적으로 시행할때 실행후 exit 하기 위한 문장
        sleep (10)  # 2분 대기후
        sys.exit()      

            # ------------------------------------------------------------------------------ #
        
    def __del__(self):
        self.obj8537.Clear()

    # 전략리스트 조회
    def listMyStrategy(self):
                
        try:            
            self.obj8537.requestList('나의',self)  # 결과는 StgList에 저장함
        except Exception as ex:
            print ('CREON API ERROR : ', ex)
            BUSbot.sendMessage('CREON API ERROR')
            sys.exit()
            
        for i in range(len(self.stgName)):
            try:
                item = self.StgList[self.stgName[i]]
                self.id.append(item['ID'])
                name = item['전략명']
            except Exception as ex:
                print (self.stgName[i] + '-전략명 없음', ex)
                BUSbot.sendMessage(self.stgName[i] + '-전략명 없음')

            ret = self.obj8537.requestStgID(self.id[i], self)  # 결과는 dataStgList 저장함
            
            if ret == True:
                print('검색전략:', self.id[i], '전략명:', name, '검색종목수:', len(self.dataStgList))
            

                #텔레그램 봇으로 메세지보내기
                BUSbot.sendMessage('전략명 : ' + name)
                
                columns = ['index','code','종목명']
                df = DataFrame.from_records(self.dataStgList, columns=columns)
                df = df.set_index(columns[0])      
                message = tabulate(df, headers='keys', tablefmt='simple')
                message = "<pre>{}</pre>".format(message)
                BUSbot.sendMessage2html(message)


    # 종목 200개이하 확인
    def total_check(self):
        if (len(self.dataStgList) >= 200):
            print('검색종목이 200 을 초과할 경우 실시간 감시 불가 ')
            return False
        return True
    
    def monitor_stg(self):
        for i in range(len(self.id)):
            try:
                ret = self.obj8537.requestMonitorID(self.id[i], self)  # self.monID 값 설정함
                if ret == True:
                    self.obj8537.requestStgControl(self.id[i], self.monID, True)
            except Exception as ex:
                
                print ('전략감시 ERROR : ', ex)
                BUSbot.sendMessage('전략감시 ERROR')
      
        return


# In[4]:


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()

