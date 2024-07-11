
import asyncio
from ib_insync import *
import datetime as dt

import pandas as pd
import sys
import pickle


ib = IB()

ib.connect("127.0.0.1", 7497, clientId=22) 



ticker='SPY'
quantity=1
account_number='DU8663689'
exchange='SMART'
currency='USD'
trading_class='SPY'
strike_limit=5


current_time=dt.datetime.now()
first_trade_flag=0
short_price=0.10
long_price=0.02
start_hour,start_min=19,1
end_hour,end_min=1,15
all_option_contract={}
shortlist_option={}



# Check if the file exists
try:
    order_filled_dataframe=pd.read_csv('order_filled_list.csv')
    order_filled_dataframe.set_index('time',inplace=True)

except:
    column_names = ['time','ticker','price','action','type1','stop_price']
    order_filled_dataframe = pd.DataFrame(columns=column_names)
    order_filled_dataframe.set_index('time',inplace=True)




contract1=ib.qualifyContracts(Stock(ticker,exchange,currency))[0]
print(contract1)


chains = ib.reqSecDefOptParams(contract1.symbol, '', contract1.secType, contract1.conId)

df1=util.df(chains)
print(df1)
df1.to_csv(f'{ticker}_chain.csv')


current_expiry=df1[(df1['exchange']==contract1.exchange) & (df1['tradingClass']==trading_class)]['expirations'].iloc[0][0]
print(current_expiry)
strike_list=df1[(df1['exchange']==contract1.exchange) & (df1['tradingClass']==trading_class)]['strikes'].iloc[0]
print(strike_list)




current_price=ib.reqTickers(contract1)[0].last

print(current_price)
new_stike_list=[i for i in strike_list if ((current_price-strike_limit )< i and i<(current_price+strike_limit))]

print('shortlisted strikes')
print(new_stike_list)

for strike in new_stike_list:
    call_option_contract=Contract(symbol=contract1.symbol,secType='OPT',exchange=contract1.exchange,lastTradeDateOrContractMonth=current_expiry, strike=strike, right='C')
    put_option_contract=Contract(symbol=contract1.symbol,secType='OPT',exchange=contract1.exchange,lastTradeDateOrContractMonth=current_expiry, strike=strike, right='P')
    # all_option_contract[call_option_contract.localSymbol]=call_option_contract
    # all_option_contract[put_option_contract.localSymbol]=put_option_contract

    c=ib.qualifyContracts(call_option_contract)
    p=ib.qualifyContracts(put_option_contract)
    if c:
        all_option_contract[c[0].localSymbol]=c[0]
    if p:
        all_option_contract[p[0].localSymbol]=p[0]

print(all_option_contract)



df=pd.DataFrame(columns=['name','times','price','oi','volume','iv','delta','gamma','vega','theta','cont_right'])
df['name']=all_option_contract.keys()
df.set_index('name',inplace=True)
print(df)


import xlwings as xw
wb = xw.Book('Data.xlsx')
sheet = wb.sheets['Sheet1']

def pending_tick_handler(t):
    t=list(t)[0]
    times=t.time.replace(tzinfo=dt.timezone.utc).astimezone(tz=None)
    name=t.contract.localSymbol 
    price=t.last if t.last else 0
    volume=t.volume if t.volume else 0
    cont_right=t.contract.right
    oi=t.callOpenInterest+t.putOpenInterest if t.callOpenInterest+t.putOpenInterest else 0

    if t.modelGreeks:
        iv=t.modelGreeks.impliedVol if t.modelGreeks.impliedVol else 0
        delta=t.modelGreeks.delta if t.modelGreeks.delta else 0  
        gamma=t.modelGreeks.gamma if t.modelGreeks.gamma else 0
        vega=t.modelGreeks.vega if t.modelGreeks.vega else 0
        theta=t.modelGreeks.theta if t.modelGreeks.theta else 0
        
    else:
         iv,delta,gamma,vega,theta=(0,0,0,0,0)

    l=[times,price,oi,volume,iv,delta,gamma,vega,theta,cont_right]

    if name:
            #updating dataframe
            df.loc[name] = l
            print(df)





for i,j  in all_option_contract.items():
    cont=j
    print(cont)
    ib.reqMktData(cont, "100, 101, 104", False, False)
    ib.sleep(2)  
    ib.reqMarketDataType(1)
    ib.pendingTickersEvent += pending_tick_handler