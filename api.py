import requests
import pandas as pd

class NSEIndia:
    # All the market segments and indices
    pre_market_keys = ['NIFTY', 'BANKNIFTY', 'SME', 'FO', 'OTHERS', 'ALL']

    live_market_keys = ['NIFTY 50', 'NIFTY NEXT 50', 'NIFTY MIDCAP 50', 'NIFTY MIDCAP 100', 'NIFTY MIDCAP 150', 
                        'NIFTY SMALLCAP 50', 'NIFTY SMALLCAP 100', 'NIFTY SMALLCAP 250', 'NIFTY MIDSMALLCAP 400', 
                        'NIFTY 100', 'NIFTY 200', 'NIFTY500 MULTICAP 50:25:25', 'NIFTY LARGEMIDCAP 250', 'NIFTY AUTO', 
                        'NIFTY BANK', 'NIFTY ENERGY', 'NIFTY FINANCIAL SERVICES', 'NIFTY FINANCIAL SERVICES 25/50', 
                        'NIFTY FMCG', 'NIFTY IT', 'NIFTY MEDIA', 'NIFTY METAL', 'NIFTY PHARMA', 'NIFTY PSU BANK', 'NIFTY REALTY', 
                        'NIFTY PRIVATE BANK', 'NIFTY HEALTHCARE INDEX', 'NIFTY CONSUMER DURABLES', 'NIFTY OIL & GAS', 
                        'NIFTY COMMODITIES', 'NIFTY INDIA CONSUMPTION', 'NIFTY CPSE', 'NIFTY INFRASTRUCTURE', 'NIFTY MNC', 
                        'NIFTY GROWTH SECTORS 15', 'NIFTY PSE', 'NIFTY SERVICES SECTOR', 'NIFTY100 LIQUID 15', 'NIFTY MIDCAP LIQUID 15', 
                        'NIFTY DIVIDEND OPPORTUNITIES 50', 'NIFTY50 VALUE 20', 'NIFTY100 QUALITY 30', 'NIFTY50 EQUAL WEIGHT', 
                        'NIFTY100 EQUAL WEIGHT', 'NIFTY100 LOW VOLATILITY 30', 'NIFTY ALPHA 50', 'NIFTY200 QUALITY 30', 
                        'NIFTY ALPHA LOW-VOLATILITY 30', 'NIFTY200 MOMENTUM 30', 'Securities in F&O', 'Permitted to Trade']
    
    holiday_keys = ['clearing', 'trading']


    def __init__(self):
        self.headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36'}
        self.session = requests.Session()
        self.session.get('http://nseindia.com', headers=self.headers)

    # NSE Pre-market data API section
    def NsePreMarketData(self, key):
        try:
            data = self.session.get(f"https://www.nseindia.com/api/market-data-pre-open?key={key}", 
                    headers=self.headers).json()['data']
        except:
            pass
        new_data = []
        for i in data:
            new_data.append(i['metadata'])
        df = pd.DataFrame(new_data)
        return df 

    # NSE Live-market data API section
    def NseLiveMarketData(self, key, symbol_list):
        try:
            data = self.session.get(f"https://www.nseindia.com/api/equity-stockIndices?index={key.upper().replace(' ','%20').replace('&','%26')}",
                    headers=self.headers).json()['data'] 
                    # Use of "replace(' ','%20').replace('&','%26')" -> In live market space is replaced by %20, & is replaced by %26
       
            df = pd.DataFrame(data)
            df = df.drop(['meta'], axis=1)
            if symbol_list:
                return list(df['symbol'])
            else:
                return df
        except:
            pass

    # NSE market holiday API section
    def NseHoliday(self, key):
        try:
            data = self.session.get(f'https://www.nseindia.com/api/holiday-master?type={key}', headers = self.headers).json()
        except:
            pass
        df = pd.DataFrame(list(data.values())[0])
        return df

class NSEIndia2:
    def __init__(self):
        try:
            self.headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36'}
            self.session = requests.Session()
            self.session.get('http://nseindia.com', headers=self.headers)
        except:
            pass

    # NSE Option-chain data API section
    def GetOptionChainData(self, symbol, indices=False):
        try:
            if not indices:
                url = 'https://www.nseindia.com/api/option-chain-equities?symbol=' + symbol
            else:
                url = 'https://www.nseindia.com/api/option-chain-indices?symbol=' + symbol
        except:
            pass
        data = self.session.get(url, headers=self.headers).json()["records"]["data"]
            
        df = []
        for i in data: 
            for keys, values in i.items():
                if keys == 'CE' or keys == 'PE':
                    info = values
                    info['instrumentType'] = keys
                    df.append(info)
        df1 = pd.DataFrame(df)
        return pd.DataFrame(df1)

