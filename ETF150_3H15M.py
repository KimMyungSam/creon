import win32com.client
from time import sleep
import sys
import CreonAPI
import ChatBotModel

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

def etf150_3h15m_buy():
    # 텔레그램 봇 생성하기
    BUSbot = ChatBotModel.Bot2ndBUS()

    # 100일간 일별 데이터 구하기, etf150_data = get_daily_price(code)
    code = 'A233740'  # ETF 코스닥150 레버리지
    ETF150 = CreonAPI.creon_func(code)

    try:
        # etf150_data = get_daily_price(code)  # 100일간 데이터 구하기
        ret = ETF150.get_daily_price()
        if ret[0] == False:
            raise get_daily_price_error()
        else:
            etf150_data = ret[1]

        ret = ETF150.get_current_price()
        # 리턴값이 int라서 if ret == False에서 비교시 type 에러가 발생함.
        if ret[0] == False:
            raise get_current_price_error()
        else:
            current_price = ret[1]
    except get_daily_price_error:
        BUSbot.sendMessage('get_daily_price : 크레온 API 연동에 문제가 발생하였습니다.')
        sys.exit()
    except get_current_price_error:
        BUSbot.sendMessage('get_current_price : 크레온 API 연동에 문제가 발생하였습니다.')
        sys.exit()
    except Exception as ex:
        print ('CREON API ERROR : ', ex)
        BUSbot.sendMessage('CREON API ERROR')
        sys.exit()

    # atr 14일 구하기, atr14 = get_atr(etf150_data)
    atr14 = ETF150.get_atr(etf150_data)
    etf150_data.insert(len(etf150_data.columns),"ATR14", atr14)

    # RISK MANAGEMENT
    one_trading_risk = 2  # 2%
    equity = 1000000  # 자본금 백만원
    one_trading_risk_price = equity * one_trading_risk / 100
    number_of_items = 1  # 종목수
    etf150_data = etf150_data.reset_index()
    ATR = etf150_data.loc[len(etf150_data)-1,'ATR14']
    stop_loss_price = current_price - (ATR * 2)  # 2N으로 계산
    one_trading_risk_price_each = one_trading_risk_price / number_of_items
    purchase_quantity = int(round(one_trading_risk_price_each / (ATR * 2)))  # 2N으로 계산
    purchase_amount = purchase_quantity * current_price
    ###########################################

    # ma 일자별 구하기
    ma = []
    for i in range(3,21):
        ma.append(ETF150.get_ma(etf150_data, i))

    for i in range(0,18):
        etf150_data.insert(len(etf150_data.columns),"ma"+str(i+3), ma[i])

    # 분단위 트랜드를 파악하여 매수 타이밍을 결정하가
    # 3시 15분에 매수할지 30분에 매수할지
    # 양봉일때는 추세가 살아 있을때 매수하고
    # 음봉일때는 추세가 죽어 있으면 종가일때 비교하여 매수하고
    import numpy as np

    kind = ""  # 그날의 매수유형 purchase, buy등 구분 변수

    min_data = ETF150.get_min_a_day(20) #1분단위 20개 close 데이타 가져오기
    coefficients, residuals, _, _, _ = np.polyfit(range(len(min_data)),min_data,1,full=True)
    mse = residuals[0]/(len(min_data))
    nrmse = np.sqrt(mse)/(max(min_data) - min(min_data))
    #print('Slope ' + str(coefficients[0]))
    #print('NRMSE: ' + str(nrmse))

    # 3시 15분에 추세파악하여 강력한 상승추세일때 그날 모두 매수
    # 시간 조건을 추가


    if coefficients[0] < 0:
        minute_buy = False
    elif coefficients[0] > 1 :
        minute_buy = True

    try:

        if minute_buy == True:
             # 현금매수하기, buy_code(code, purchase_quantity, buy_price)
            code = 'A233740'  # ETF 코스닥150 레버리지
            ETF150_deal = CreonAPI.buy_code(code)

            ret = ETF150_deal.buy(purchase_quantity, current_price)  # 매수 수량, 매수가격=현재가격
            if ret == False:
                raise BuyError()
            else:
                BUSbot.sendMessage('3시15분에 상승Trend로 크레온 API를 통해 매수하였습니다.')
            kind = "purchase"

    except BuyError:
        print ("매수시 크레온 API 연동에 문제가 발생하였습니다")  #telegram bot으로 통보하기
        BUSbot.sendMessage('매수시 크레온 API 연동에 문제가 발생하였습니다.')
        sys.exit()
    except Exception as ex:
        print ('CREON API ERROR : ', ex)
        BUSbot.sendMessage('CREON API ERROR')
        sys.exit()

    # 텔레그램 보내기
    import matplotlib.pyplot as plt

    fig = plt.figure(figsize=(20,10))
    l = fig.add_subplot(1,2,1)
    r = fig.add_subplot(1,2,2)

    l.plot(etf150_data[95:100].Close,'o',color='black', linestyle='dashed',markersize=12)
    l.plot(etf150_data[95:100].Open,'o',color='red', linestyle='dashed',markersize=12)
    l.plot(etf150_data[95:100].ma5)
    l.plot(etf150_data[95:100].ma10)
    l.plot(etf150_data[95:100].ma15)
    l.plot(etf150_data[95:100].ma20)
    l.legend(loc='best')
    l.grid(True)

    r.text(0.0, 0.9, "1. Equity : " + str(equity), size=20,ha="left", va="center")
    r.text(0.0, 0.8, "2. one trading risk(%) : "+ str(one_trading_risk),size=20,ha="left", va="center")
    r.text(0.0, 0.7, "3. Number of items : "+ str(number_of_items),size=20,ha="left", va="center")
    r.text(0.0, 0.6, "4. ATR(2N) Price : "+ str(ATR * 2),size=20,ha="left", va="center")
    r.text(0.0, 0.5, "5. Current Price : "+ str(current_price),size=20,ha="left", va="center")
    r.text(0.0, 0.4, "6. Stop Loss : "+ str(stop_loss_price),size=20,ha="left", va="center")

    if kind == "purchase":
        r.text(0.0, 0.3, "7. Purchase quantity : "+ str(purchase_quantity),size=20,ha="left", va="center")
        r.text(0.0, 0.2, "8. Purchase amount : "+ str(purchase_amount),size=20,ha="left", va="center")
    elif kind == "buy":
        r.text(0.0, 0.3, "7. Buy quantity : "+ str(BUY_each),size=20,ha="left", va="center")
        r.text(0.0, 0.2, "8. Buy amount : "+ str(BUY_amount),size=20,ha="left", va="center")
    else:
        r.text(0.0, 0.3, "== I do not trade today ==",size=30, ha="left", va="center")

    r.plot
    # r.plot(x,y)
    plt.show()
    fig.savefig("etf150.png")

    # 매매 현황을 이미지 파일로 전송
    BUSbot.sendPhoto(open('etf150.png', 'rb'))

    return True
