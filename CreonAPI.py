import pandas as pd
import win32com.client

class creon_func:
    def __init__(self, code):
        self.code = code

        # 연결 여부 체크
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = self.objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

    # API를 통한 일별 데이터 조회
    def get_daily_price(self):
        self.stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")

        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.stock_chart.GetDibStatus()
        rqRet = self.stock_chart.GetDibMsg1()
        print("get_daily_price 통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return (False,) # 튜플은 1개 요소일때 콤마(,) 붙여야함.
                            # 정상적인 리턴일때 2개임으로 false도 듀플로 리턴함


        self.stock_chart.SetInputValue(0, self.code)                          # 종목 코드
        self.stock_chart.SetInputValue(1, ord('2'))                      # 요청 구분 '1': 기간, '2': 개수
        self.stock_chart.SetInputValue(4, 100)                           # 요청 데이터의 개수
        self.stock_chart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 13])        # 요청 내용, 0-day, 2-open, 3,-high..
        self.stock_chart.SetInputValue(6, ord('D'))                      # 분봉 데이터
        self.stock_chart.SetInputValue(9, '1')                           # 수정 주가
        self.stock_chart.BlockRequest()

        count = self.stock_chart.GetHeaderValue(3)

        price_list = []
        for i in range(count):
            day = self.stock_chart.GetDataValue(0, i)  # 위 SetInputValue(5, [0, 2, 3, 4, 5, 8, 13]) 결과를 순서대로 get
            open = self.stock_chart.GetDataValue(1, i)
            high = self.stock_chart.GetDataValue(2, i)
            low = self.stock_chart.GetDataValue(3, i)
            close = self.stock_chart.GetDataValue(4, i)
            volume = self.stock_chart.GetDataValue(5, i)

            price_list.append([day, open, high, low, close, volume])

        labels = ['Day','Open','High','Low','Close','Volume']
        df = pd.DataFrame.from_records(price_list, columns=labels)
        # API에서 데이터를 가져오면 최신데이타가 제일 먼저임.
        # 과거로 부터 현재까지 시계열로 데이타 분석시 재 정렬하여 사용하여야 함.
        df = df.sort_values(by=['Day'], ascending=True)  # 날짜 기준으로 오름차순 정렬
        #df['Day'] = pd.to_datetime(df['Day'])
        value = df.set_index('Day')

        return (True, value)

    # 14일 ATR구하기
    def get_atr(self, raw):
        TR1 = raw['High'] - raw['Low']
        TR2 = abs(raw['Close'].shift(1) - raw['High'])
        TR3 = abs(raw['Close'].shift(1) - raw['Low'])

        df = pd.concat([TR1, TR2, TR3], axis=1)
        tr = df.max(axis=1)
        atr13 = tr.rolling(window=13).mean()
        atr14 = ((atr13.shift(1) * 13 + tr) / 14)

        return round(atr14)

    # MA 구하기
    def get_ma(self, raw, day):
        ma = raw['Close'].rolling(window=day).mean()

        return round(ma)

    # 이동평균과 현재 가격을 비교 매수할 %(비율) 점수 구하기
    def get_score(self, ma):
        ma = ma.iloc[-1]

        # 현재가 객체 구하기
        objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        objStockMst.SetInputValue(0, self.code)   #종목 코드 - 코스닥150 레버리지
        objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objStockMst.GetDibStatus()
        rqRet = objStockMst.GetDibMsg1()
        print("get_score 통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return (False,)

        # 현재가 정보 조회
        #code = objStockMst.GetHeaderValue(0)  #종목코드
        #name= objStockMst.GetHeaderValue(1)  # 종목명
        #time= objStockMst.GetHeaderValue(4)  # 시간
        cprice= objStockMst.GetHeaderValue(11) # 종가, 현재가격
        #diff= objStockMst.GetHeaderValue(12)  # 대비
        #open= objStockMst.GetHeaderValue(13)  # 시가
        #high= objStockMst.GetHeaderValue(14)  # 고가
        #low= objStockMst.GetHeaderValue(15)   # 저가
        #offer = objStockMst.GetHeaderValue(16)  #매도호가
        #bid = objStockMst.GetHeaderValue(17)   #매수호가
        #vol= objStockMst.GetHeaderValue(18)   #거래량
        #vol_value= objStockMst.GetHeaderValue(19)  #거래대금

        score = 0
        for i in range(6,len(ma)):  # ma3 loc가 6 이고 ma20이 끝 행
            if cprice > ma[i]:
                score += 1

        return (True, (score/len(ma)))


    # 양봉, 음봉 파악하여 매수 하기

    def get_market_trend(self, data):

        # 현재가 객체 구하기
        objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        objStockMst.SetInputValue(0, self.code)   #종목 코드
        objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objStockMst.GetDibStatus()
        rqRet = objStockMst.GetDibMsg1()
        print("get_market_trend 통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return (False,)

        # 현재가 정보 조회
        #code = objStockMst.GetHeaderValue(0)  #종목코드
        #name= objStockMst.GetHeaderValue(1)  # 종목명
        #time= objStockMst.GetHeaderValue(4)  # 시간
        cprice= objStockMst.GetHeaderValue(11) # 종가, 현재가격
        #diff= objStockMst.GetHeaderValue(12)  # 대비
        open= objStockMst.GetHeaderValue(13)  # 시가
        high= objStockMst.GetHeaderValue(14)  # 고가
        low= objStockMst.GetHeaderValue(15)   # 저가
        #offer = objStockMst.GetHeaderValue(16)  #매도호가
        #bid = objStockMst.GetHeaderValue(17)   #매수호가
        #vol= objStockMst.GetHeaderValue(18)   #거래량
        #vol_value= objStockMst.GetHeaderValue(19)  #거래대금

        if cprice > open:  # 양봉이면 매수
            state = "UP"
        elif cprice < open and cprice > data.iloc[-2][3]:  # 음봉이지만 전날 종가보다 상승하고 있으면 매수, data.iloc[-2][3] -- 전날 close 가격
            state = "UP"
        else:
            state = "DOWN"

        return (True, state, cprice)

    # 현재가 구하기
    def get_current_price(self):

        # 현재가 객체 구하기
        objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        objStockMst.SetInputValue(0, self.code)   #종목 코드
        objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objStockMst.GetDibStatus()
        rqRet = objStockMst.GetDibMsg1()
        print("get_current_price 통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return (False,)

        # 현재가 정보 조회
        #code = objStockMst.GetHeaderValue(0)  #종목코드
        #name= objStockMst.GetHeaderValue(1)  # 종목명
        #time= objStockMst.GetHeaderValue(4)  # 시간
        cprice= objStockMst.GetHeaderValue(11) # 종가, 현재가격
        #diff= objStockMst.GetHeaderValue(12)  # 대비
        #open= objStockMst.GetHeaderValue(13)  # 시가
        #high= objStockMst.GetHeaderValue(14)  # 고가
        #low= objStockMst.GetHeaderValue(15)   # 저가
        #offer = objStockMst.GetHeaderValue(16)  #매도호가
        #bid = objStockMst.GetHeaderValue(17)   #매수호가
        #vol= objStockMst.GetHeaderValue(18)   #거래량
        #vol_value= objStockMst.GetHeaderValue(19)  #거래대금

        return (True, cprice)

    # 1일치 분데이타 구하기
    def get_min_a_day(self, each):
        import win32com.client
        self.stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")

        self.stock_chart.SetInputValue(0, self.code)                     # 종목 코드
        self.stock_chart.SetInputValue(1, ord('2'))                      # 요청 구분 '1': 기간, '2': 개수
        #self.stock_chart.SetInputValue(2, 20180531)                      # 시작/끝 날짜를 같이 주어야 하루치 분데이터 조회가능능
        #self.stock_chart.SetInputValue(3, 20180531)
        self.stock_chart.SetInputValue(4, each)                           # 요청 데이터의 개수
        self.stock_chart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 13])        # 요청 내용, 0-day, 2-open, 3,-high..
        self.stock_chart.SetInputValue(6, ord('m'))                      # 분봉 데이터
        self.stock_chart.SetInputValue(9, '1')                           # 수정 주가
        self.stock_chart.BlockRequest()

        count = self.stock_chart.GetHeaderValue(3)

        val_list =[]
        for i in range(count):
            #day = self.stock_chart.GetDataValue(0, i)  # 위 SetInputValue(5, [0, 2, 3, 4, 5, 8, 13]) 결과를 순서대로 get
            #open = self.stock_chart.GetDataValue(1, i)
            #high = self.stock_chart.GetDataValue(2, i)
            #low = self.stock_chart.GetDataValue(3, i)
            close = self.stock_chart.GetDataValue(4, i)
            #volume = self.stock_chart.GetDataValue(5, i)
            #cap = self.stock_chart.GetDataValue(6, i)

            val_list.append(close)
        val_list.reverse()  # API데이타가 최신데이터가 먼저옴으로 과거데이타에서 최근으로 시계열 순서를 바꿈
        return (val_list)

class order:
    # 주식 잔고 조회
    def rq6033(self):
        # 통신 OBJECT 기본 세팅
        self.objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
        initCheck = self.objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            return (False,)

        # 계좌번호
        self.acc = self.objTrade.AccountNumber[0]  # 계좌번호
        accFlag = self.objTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        print(self.acc, accFlag[0])

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, self.acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  #  요청 건수(최대 50)
        # 6033 통신처리
        self.objRq.BlockRequest()

        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return (False,)

        cnt = self.objRq.GetHeaderValue(7)  # 7 - (long) 수신개수

        jangoData = {}
        for i in range(cnt):
            item = {}
            code = self.objRq.GetDataValue(12, i)  # 종목코드
            item['종목코드'] = code
            item['종목명'] = self.objRq.GetDataValue(0, i)  # 종목명
            item['현금신용'] = self.objRq.GetDataValue(1, i)  # 신용구분
            print(code, '현금신용', item['현금신용'])
            item['대출일'] = self.objRq.GetDataValue(2, i)  # 대출일
            item['잔고수량'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
            item['매도가능'] = self.objRq.GetDataValue(15, i)
            item['장부가'] = self.objRq.GetDataValue(17, i)  # 체결장부단가
            # 매입금액 = 장부가 * 잔고수량
            item['매입금액'] = item['장부가'] * item['잔고수량']

            # 잔고 추가
            jangoData[code] = item

            if len(jangoData) >= 200:  # 최대 200 종목만,
                break

        return (True, jangoData)

    def modifyOrder(self, ordernum, price):
        # 주식 정정 주문
        print("정정주문", ordernum, self.code, price)
        self.objModifyOrder.SetInputValue(1, ordernum)  #  원주문 번호 - 정정을 하려는 주문 번호
        self.objModifyOrder.SetInputValue(2, self.acc)  # 상품구분 - 주식 상품 중 첫번째
        self.objModifyOrder.SetInputValue(3, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objModifyOrder.SetInputValue(4, self.code)  # 종목코드
        self.objModifyOrder.SetInputValue(5, 0)  # 정정 수량, 0 이면 잔량 정정임
        self.objModifyOrder.SetInputValue(6, price)  #  정정주문단가

        # 정정주문 요청
        self.objModifyOrder.BlockRequest()

        rqStatus = self.objModifyOrder.GetDibStatus()
        rqRet = self.objModifyOrder.GetDibMsg1()
        print("modifyOrder 통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
        else:
            return True
        # 새로운 주문 번호 구한다.
        self.orderNum = self.objModifyOrder.GetHeaderValue(7)

    def cancelOrder(self, ordernum):
        # 주식 취소 주문
        print("취소주문", ordernum, self.code)
        self.objCancelOrder.SetInputValue(1, ordernum)  #  원주문 번호 - 정정을 하려는 주문 번호
        self.objCancelOrder.SetInputValue(2, self.acc)  # 상품구분 - 주식 상품 중 첫번째
        self.objCancelOrder.SetInputValue(3, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objCancelOrder.SetInputValue(4, self.code)  # 종목코드
        self.objCancelOrder.SetInputValue(5, 0)  # 정정 수량, 0 이면 잔량 취소임

        # 취소주문 요청
        self.objCancelOrder.BlockRequest()

        rqStatus = self.objCancelOrder.GetDibStatus()
        rqRet = self.objCancelOrder.GetDibMsg1()
        print("cancelOrder 통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
        else:
            return True

# 주식 매수 주문
# 주식 잔고 조회 def rq6033(self)상속 # 정정/취소주문 상속
class buy_code(order):
    def __init__(self,code):
        self.code = code

        # 연결 여부 체크
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = self.objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        # 주문 초기화
        self.objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
        initCheck = self.objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            return False

        self.acc = self.objTrade.AccountNumber[0] #계좌번호
        self.accFlag = self.objTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        print(self.acc, self.accFlag[0])

        # 매수/정정/취소 주문 object 생성
        self.objBuyOrder = win32com.client.Dispatch("CpTrade.CpTd0311")     # 매수
        #self.objModifyOrder = win32com.client.Dispatch("CpTrade.CpTd0313")  # 정정
        #self.objCancelOrder = win32com.client.Dispatch("CpTrade.CpTd0314")  # 취소
        self.orderNum = 0 # 주문 번호

    def buy(self, buy_each, buy_price):
        # 주식 매수 주문
        self.objBuyOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
        self.objBuyOrder.SetInputValue(0, "2")   # 2: 매수
        self.objBuyOrder.SetInputValue(1, self.acc )   #  계좌번호
        self.objBuyOrder.SetInputValue(2, self.accFlag[0])   # 상품구분 - 주식 상품 중 첫번째
        self.objBuyOrder.SetInputValue(3, self.code)   # 종목코드
        self.objBuyOrder.SetInputValue(4, buy_each)   # 매수수량
        self.objBuyOrder.SetInputValue(5, buy_price)   # 주문단가
        self.objBuyOrder.SetInputValue(7, "0")   # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objBuyOrder.SetInputValue(8, "01")   # 주문호가 구분코드 - 01: 보통

        # 매수 주문 요청
        self.objBuyOrder.BlockRequest()

        rqStatus = self.objBuyOrder.GetDibStatus()
        rqRet = self.objBuyOrder.GetDibMsg1()
        print("buy_code 통신상태", rqStatus, rqRet)
        if rqStatus != 0:  # fail일때
            return False
        else:
            return True
        # 주의: 매수 주문에  대한 구체적인 처리는 cpconclution 으로 파악해야 한다.

# 주식 매도 주문
# 주식 잔고 조회 def rq6033(self)상속 # 정정/취소주문 상속
class sell_code(order):
    def __init__(self,code):
        self.code = code

        # 연결 여부 체크
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = self.objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        # 주문 초기화
        self.objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
        initCheck = self.objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            return False

        # 주식 매도 주문
        self.acc = self.objTrade.AccountNumber[0] #계좌번호
        self.accFlag = self.objTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        print(self.acc, self.accFlag[0])

        # 매도/정정/취소 주문 object 생성
        self.objSellOrder = win32com.client.Dispatch("CpTrade.CpTd0311")     # 매도
        #self.objModifyOrder = win32com.client.Dispatch("CpTrade.CpTd0313")  # 정정
        #self.objCancelOrder = win32com.client.Dispatch("CpTrade.CpTd0314")  # 취소
        self.orderNum = 0 # 주문 번호

    def sell(self,buy_each, buy_price):

        self.objSellOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
        self.objSellOrder.SetInputValue(0, "1")   #  1: 매도
        self.objSellOrder.SetInputValue(1, self.acc )   #  계좌번호
        self.objSellOrder.SetInputValue(2, self.accFlag[0])   #  상품구분 - 주식 상품 중 첫번째
        self.objSellOrder.SetInputValue(3, self.code)   #  종목코드 - A003540 - 대신증권 종목
        self.objSellOrder.SetInputValue(4, 10)   #  매도수량 10주
        self.objSellOrder.SetInputValue(5, 14100)   #  주문단가  - 14,100원
        self.objSellOrder.SetInputValue(7, "0")   #  주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objSellOrder.SetInputValue(8, "01")   # 주문호가 구분코드 - 01: 보통

        # 매도 주문 요청
        self.objSellOrder.BlockRequest()

        rqStatus = self.objSellOrder.GetDibStatus()
        rqRet = self.objSellOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0: # fail일때
            return False
        else:
            return True

class BuyError(Exception):
    pass

class get_market_trend_error(Exception):
    pass

class get_score_error(Exception):
    pass

class get_current_price_error(Exception):
    pass

class get_daily_price_error(Exception):
    pass

'''
value = object.GetDataValue(Type,Index)

type에해당하는데이터를반환합니다

type: 데이터종류

0 - (string) 종목명

1 - (char)신용구분

코드 - 내용
'Y'-신용융자/유통융자

'D'-신용대주/유통대주

'B'-담보대출

'M'-매입담보대출

'P'-플러스론대출

'I'-자기융자/유통융자

2 - (string) 대출일

3 - (long)결제잔고수량

4 - (long)결제장부단가

5 - (long)전일체결수량

6 - (long)금일체결수량

7 - (long)체결잔고수량

9 - (longlong)평가금액(단위:원) - 천원미만은내림

10 - (longlong)평가손익(단위:원) - 천원미만은내림

11 - (double)수익률

12 - (string) 종목코드

13 - (char)주문구분

15 - (long)매도가능수량

16 - (string) 만기일

17 - (double) 체결장부단가

18 - (longlong) 손익단가

반환값: 데이터종류의 index번째 data
'''

'''
value = object.GetHeaderValue(type)

type에해당하는헤더데이터를반환합니다

type: 데이터종류

0 - (string) 계좌명

1 - (long) 결제잔고수량

2 - (long)체결잔고수량

3 - (longlong)평가금액(단위:원)

4 - (longlong)평가손익(단위:원)

5 - 사용하지않음

6 - (longlong)대출금액(단위:원)

7 - (long) 수신개수

8 - (double) 수익율

9 - (longlong) D+2 예상예수금

10 - (longlong) 대주평가금액

11 - (longlong) 잔고평가금액

12 - (longlong) 대주금액

반환값: 데이터종류에해당하는값
'''
