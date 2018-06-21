import win32com.client
from time import sleep
import sys
import ChatBotModel
import CreonAPI

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

def etf150_3h30m_buy():

    # if __name__ == "__main__":
    BUS = ChatBotModel.Bot2ndBUS()  # 텔레그램 봇 기동

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
        BUS.sendMessage('get_daily_price 크레온 API 연동에 문제가 발생하였습니다.')
        sys.exit()
    except get_current_price_error:
        BUS.sendMessage('get_current_price 크레온 API 연동에 문제가 발생하였습니다.')
        sys.exit()
    except Exception as ex:
        print ('CREON API ERROR', ex)
        BUS.sendMessage('CREON API ERROR')
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


    # In[5]:


    # 전략1. 총5회 매수, 3시 10분, 15분, 20분, 25분, 28분
    # 전략2. 3시15분에 직전 20분간 트랜드를 보고 상승일때 그날의 purchase_quantity만큼 매수
    kind = ""  # 매수유형을 정의하는 변수

    try:

        # 현재가격 기준으로 score 구하기, score = get_score(etf150_data, code)
        ret = ETF150.get_score(etf150_data)
        if ret[0] == False:
            raise get_score_error()
        else:
            score = ret[1]

        # MA 스코어와 변동성 2N을 곱하여 최종적으로 구매해야할 수량 결정
        BUY_each = round (score * purchase_quantity)
        BUY_amount = BUY_each * current_price


        # 매수조건 구하기, 리턴할때 UP/DOWN과 현재가 돌려줌 entry = get_market_trend(etf150_data)
        ret = ETF150.get_market_trend(etf150_data)  # return값으로 [0] - UP/Down, [1] - 현재가
        if ret[0] == False:
            raise get_market_trend_error()
        else:
            entry = ret

         # 현금매수하기, buy_code(code, purchase_quantity, buy_price)
        code = 'A233740'  # ETF 코스닥150 레버리지
        ETF150_deal = CreonAPI.buy_code(code)


        if entry[1] == "UP":
            ret = ETF150_deal.buy(BUY_each, entry[2])  # 매수 수량, 매수가격=현재가격
            if ret == False:
                raise BuyError()
            else:
                BUS.sendMessage('크레온 API를 통해 매수하였습니다.')
            kind = "buy"
        else:
            print ("금일 매수 신호가 발생하지 않았습니다.")
            BUS.sendMessage('금일 매수 신호가 발생하지 않았습니다.')

    except get_market_trend_error:
        BUS.sendMessage('get_market_trend 크레온 API 연동에 문제가 발생하였습니다.')
        sys.exit()
    except get_score_error:
        BUS.sendMessage('get_score 크레온 API 연동에 문제가 발생하였습니다.')
        sys.exit()
    except BuyError:
        print ("매수시 크레온 API 연동에 문제가 발생하였습니다")  #telegram bot으로 통보하기
        BUS.sendMessage('매수시 크레온 API 연동에 문제가 발생하였습니다.')
        sys.exit()
    except Exception as ex:
        print ('CREON API ERROR', ex)
        BUS.sendMessage('CREON API ERROR')
        sys.exit()

    # csv로 저장
    etf150_data.to_csv("etf150.csv", mode="w")

    # 텔레그램 보내기
    import matplotlib.pyplot as plt
    get_ipython().run_line_magic('matplotlib', 'inline')

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
    BUS.sendPhoto(open('etf150.png', 'rb'))

    return True
