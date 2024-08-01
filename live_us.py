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
end_hour,end_min=1,20
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

def store(data):
    pickle.dump(data,open(f'data{dt.date.today()}.pickle','wb'))

def load():
    return pickle.load(open(f'data{dt.date.today()}.pickle', 'rb'))

def updat_order_csv(name,price,action,type1,stop_price):
    global order_filled_dataframe
    a=[name,price,action,type1,stop_price]
    order_filled_dataframe.loc[dt.datetime.now()] = a
    order_filled_dataframe.to_csv('order_filled_list.csv')




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



print('all option contract')
print(all_option_contract)
df=pd.DataFrame(columns=['name','times','price','oi','volume','iv','delta','gamma','vega','theta','cont_right'])
df['name']=all_option_contract.keys()
df.set_index('name',inplace=True)


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
            # print(df)



#closing all open orders
async def close_all_orders():
    # logging.info('closing all open orders')
    Trades1 = await ib.reqAllOpenOrdersAsync()
    # print(Trades1)
    for trade in Trades1:
        # logging.info(order)
        a=ib.cancelOrder(trade.order)
        print(a)
        # logging.info(a)
    return 1
      

async def close_all_position():
    positions =await ib.reqPositionsAsync()  # A list of positions, according to IB
    print(positions)
    # logging.info(positions)
    for position in positions:
        # logging.info(position)
        print(position)
        n = position.contract.localSymbol
        # contract = Contract(conId=position.contract.conId,exchange='NSE',currency='INR')
        contract=position.contract
        print(contract)

        c=await ib.qualifyContractsAsync(contract)
        c=c[0]
        if position.position==0:
            continue
        if position.position > 0: # Number of active Long positions
            action1 = 'SELL' # to offset the long positions
        elif position.position < 0: # Number of active Short positions
            action1 = 'BUY' # to offset the short positions
        totalQuantity = int(abs(position.position))
        print(action1)
        print(totalQuantity)
        # logging.info(f'Flatten Position: {contract} {totalQuantity} {action}')
        order1 = MarketOrder(action1,totalQuantity)
        trade = ib.placeOrder(c, order1)
        print(trade)
        # logging.info(trade)

    return 1


async def get_nearest_cent_option(df,cent,right):
    df1=df[df['cont_right']==right]
    option_name=(df1['price'] - cent).abs().idxmin()
    return option_name




pos=ib.reqPositions()
try:
    data=load()
    print(data)
    print('data loaded from pickle file')
    shortlist_option=data
    first_trade_flag=shortlist_option['first_trade_flag']
except:
    print('no data to read')


print(pos)
# if len(pos)!=0:
#     pos_df=util.df(pos)
#     print('position inside')
#     if len(pos_df)>0:
#         pos_df=pos_df[pos_df['position']!=0]
#         if len(pos_df)>3:
#             pos_df['name']=[cont.localSymbol for cont in pos_df['contract']]
#             pos_df['right']=[cont.right for cont in pos_df['contract']]
#             shortlist_option={}
#             shortlist_option['short_call_option']={'name':pos_df[(pos_df['right']=='C') & (pos_df['position']<0)].name.iloc[0],'contract':pos_df[(pos_df['right']=='C') & (pos_df['position']<0)].contract.iloc[0],'buy_price':pos_df[(pos_df['right']=='C') & (pos_df['position']<0)].avgCost.iloc[0] }
#             shortlist_option['short_put_option']={'name':pos_df[(pos_df['right']=='P') & (pos_df['position']<0)].name.iloc[0],'contract':pos_df[(pos_df['right']=='P') & (pos_df['position']<0)].contract.iloc[0],'buy_price':pos_df[(pos_df['right']=='P') & (pos_df['position']<0)].avgCost.iloc[0] }
#             shortlist_option['long_call_option']={'name':pos_df[(pos_df['right']=='C') & (pos_df['position']>0)].name.iloc[0],'contract':pos_df[(pos_df['right']=='C') & (pos_df['position']>0)].contract.iloc[0],'buy_price':pos_df[(pos_df['right']=='C') & (pos_df['position']>0)].avgCost.iloc[0] }
#             shortlist_option['long_put_option']={'name':pos_df[(pos_df['right']=='P') & (pos_df['position']>0)].name.iloc[0],'contract':pos_df[(pos_df['right']=='P') & (pos_df['position']>0)].contract.iloc[0],'buy_price':pos_df[(pos_df['right']=='P') & (pos_df['position']>0)].avgCost.iloc[0] }
#             first_trade_flag=1
#             # ord_df=util.df(ib.reqAllOpenOrders())
#             # if len(ord_df)>0:
#             #     first_trade_flag=3
#     else:
#         print('position is empty')


def buy_condor(shortlist_option):
    for name,data in shortlist_option.items() :   
        if name.startswith("short"):
            direction='SELL'
        else:
            direction='BUY'
        contract_object=data['contract']
        order_object=MarketOrder(direction,quantity)
        pd1=ib.placeOrder(contract_object,order_object)
        shortlist_option[name]['order_placed']=True   
        updat_order_csv(contract_object.localSymbol,df.loc[contract_object.localSymbol,'price'],direction,'MKT',0) 
    return shortlist_option

async def manage_iron_condor(shortlist_option,option_leg,new_price):
    #get position
    pos=await ib.reqPositionsAsync()
    pos_df=util.df(pos)
    pos_df['name']=[cont.localSymbol for cont in pos_df['contract']]
    pos_df=pos_df[pos_df['position']!=0]
    print(pos_df)
    #close option leg
    leg_name=shortlist_option[option_leg].get('name')
    print(leg_name)
    if leg_name in pos_df['name'].to_list():
        print('inside')
        contract2=shortlist_option[option_leg].get('contract')
        order_object=MarketOrder('BUY',quantity)     
        pd1=ib.placeOrder(contract2,order_object)
        updat_order_csv(contract2.localSymbol,df.loc[contract2.localSymbol,'price'],'BUY','MKT',0)
        pd1
        print(pd1)

        #update option leg
      
        new_option=await get_nearest_cent_option(df,new_price,contract2.right)
        shortlist_option[option_leg]={'name':new_option,'contract':all_option_contract[new_option],'buy_price':df.loc[new_option,'price']}
        new_option_contract=all_option_contract[new_option]
        ord_obj=MarketOrder('SELL',quantity)
        pd1=ib.placeOrder(new_option_contract,ord_obj)
        updat_order_csv(new_option_contract.localSymbol,df.loc[new_option_contract.localSymbol,'price'],'SELL','MKT',0)
        print(pd1)

    return shortlist_option



async def change_stop_order_price(shortlist_option,option_type):

    #cancel stop order
    #place new stop order half the price
    cont=shortlist_option[option_type].get('contract')
    print(cont)
    cont=await ib.qualifyContractsAsync(cont)
    cont=cont[0]
    print(df)
    p=df.loc[cont.localSymbol,'price']
    print(p)
    sp=shortlist_option[option_type]['stop_price']

    ord1=StopOrder('BUY',quantity,sp/2)
    ib.cancelOrder(shortlist_option[option_type].get('stop_order_object'))
    updat_order_csv(shortlist_option[option_type].get('name'),df.loc[shortlist_option[option_type].get('name'),'price'],'BUY','CANCEL',shortlist_option[option_type].get('stop_price'))
    p1=ib.placeOrder(cont,ord1)
    updat_order_csv(cont.localSymbol,df.loc[cont.localSymbol,'price'],'BUY','UPDATE',sp/2)
    print(p1)
    return shortlist_option


async def stop_order_on_leg(shortlist_option,df):
    #place for call option
    print('inside stop order leg')
    cont=shortlist_option['short_call_option'].get('contract')
    print(cont)
    cont=await ib.qualifyContractsAsync(cont)
    cont=cont[0]
    print(df)
    p=df.loc[cont.localSymbol,'price']
    print(p)
    ord1=StopOrder('BUY',quantity,p*2)
    p1=ib.placeOrder(cont,ord1)
    updat_order_csv(cont.localSymbol,df.loc[cont.localSymbol,'price'],'BUY','STP',p*2)
    print(p1)
    shortlist_option['short_call_option']['stop_price']=p*2
    shortlist_option['short_call_option']['stop_order_object']=ord1

    #place for put option

    cont=shortlist_option['short_put_option'].get('contract')
    cont=await ib.qualifyContractsAsync(cont)
    cont=cont[0]
    p=df.loc[cont.localSymbol,'price']
    ord1=StopOrder('BUY',quantity,p*2)
    p1=ib.placeOrder(cont,ord1)
    updat_order_csv(cont.localSymbol,df.loc[cont.localSymbol,'price'],'BUY','STP',p*2)
    print(p1)
    shortlist_option['short_put_option']['stop_price']=p*2
    shortlist_option['short_put_option']['stop_order_object']=ord1
    
    return shortlist_option



for i,j  in all_option_contract.items():
    cont=j
    print(cont)
    ib.reqMktData(cont, "100, 101, 104", False, False)
    ib.sleep(2)  
    ib.reqMarketDataType(1)
    ib.pendingTickersEvent += pending_tick_handler


async def main():
    global shortlist_option
    global first_trade_flag

    while True:

        await asyncio.sleep(1)
        # sheet['A2'].value = df
        # print(df)
        # print(dt.datetime.now())

        if first_trade_flag==0 and dt.datetime.now()>dt.datetime(current_time.year,current_time.month,current_time.day,start_hour,start_min):
            print('taking trade')
            

            short_call_option=await get_nearest_cent_option(df,short_price,'C')
            
            shortlist_option['short_call_option']={'name':short_call_option,'contract':all_option_contract[short_call_option],'buy_price':df.loc[short_call_option,'price']}

            short_put_option=await get_nearest_cent_option(df,short_price,'P')
            shortlist_option['short_put_option']={'name':short_put_option,'contract':all_option_contract[short_put_option],'buy_price':df.loc[short_put_option,'price']}

            long_call_option=await get_nearest_cent_option(df,long_price,'C')
            shortlist_option['long_call_option']={'name':long_call_option,'contract':all_option_contract[long_call_option],'buy_price':df.loc[long_call_option,'price']}

            long_put_option=await get_nearest_cent_option(df,long_price,'P')
            shortlist_option['long_put_option']={'name':long_put_option,'contract':all_option_contract[long_put_option],'buy_price':df.loc[long_put_option,'price']}


            print('short_call_option',short_call_option,df.loc[short_call_option,'price'])
            print('short_put_option',short_put_option,df.loc[short_put_option,'price'])
            print("long_call_option",long_call_option,df.loc[long_call_option,'price'])
            print("long_put_option",long_put_option,df.loc[long_put_option,'price'])

            first_trade_flag=1
            shortlist_option=buy_condor(shortlist_option)
            shortlist_option['first_trade_flag']=1
            store(shortlist_option)




        if first_trade_flag==1:
            #check if placed order price doubled
            print('order placed')
            print(shortlist_option) 
            if shortlist_option['short_call_option'].get('buy_price')*2<df.loc[shortlist_option['short_call_option'].get('name'),'price']:
                shortlist_option=await manage_iron_condor(shortlist_option,'short_put_option',shortlist_option['short_call_option'].get('buy_price')*2)
                first_trade_flag=2
                shortlist_option['first_trade_flag']=2
                store(shortlist_option)
            if shortlist_option['short_put_option'].get('buy_price')*2<df.loc[shortlist_option['short_put_option'].get('name'),'price']:
                shortlist_option=await manage_iron_condor(shortlist_option,'short_call_option',shortlist_option['short_put_option'].get('buy_price')*2)
                first_trade_flag=2
                shortlist_option['first_trade_flag']=2
                store(shortlist_option)

        if first_trade_flag==2:
            #place stop order on both legs
            print('first trade 2')
            shortlist_option=await stop_order_on_leg(shortlist_option,df)
            first_trade_flag=3 
            shortlist_option['first_trade_flag']=3
            store(shortlist_option)
 


        if first_trade_flag==3:
            print('first trade 3')
            print(shortlist_option)
       
            if df.loc[shortlist_option['short_call_option'].get('name'),'price']>shortlist_option['short_call_option']['stop_price']:
                updat_order_csv(shortlist_option['short_call_option'].get('name'),df.loc[shortlist_option['short_call_option'].get('name'),'price'],'BUY','STPFill',0)
                shortlist_option=await change_stop_order_price(shortlist_option,'short_put_option')
                first_trade_flag=4
                shortlist_option['first_trade_flag']=4
                store(shortlist_option)
            elif df.loc[shortlist_option['short_put_option'].get('name'),'price']>shortlist_option['short_put_option']['stop_price']:
                updat_order_csv(shortlist_option['short_put_option'].get('name'),df.loc[shortlist_option['short_put_option'].get('name'),'price'],'BUY','STPFill',0)
                shortlist_option=await change_stop_order_price(shortlist_option,'short_call_option')
                first_trade_flag=4
                shortlist_option['first_trade_flag']=4
                store(shortlist_option)
   
                
        if first_trade_flag==4:
            print('first flag is 4')
            print(shortlist_option)


        if dt.datetime.now()>dt.datetime(current_time.year,current_time.month,current_time.day,end_hour,end_min):
            await close_all_orders()
            await close_all_position()
            ib.disconnect()
            sys.exit()



ib.run(main())










