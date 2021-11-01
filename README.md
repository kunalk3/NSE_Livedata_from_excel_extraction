# NSE Livedata from excel
We can fetch the NSE live data using API from the excel itself, built python exe based application which is responsible to fetch live market data from NSE India such as LiveMarket, PreMarket, NseHolidays, NseFNO, Symboldata, and so on... Here, we automate excel itself for live data extraction using xlwings. Here demonstrated with python tk interface and built exe based application _NSEApplication.exe_.

Currenly script/ application collect NSE data as __Live Market Data, Pre Market Data, Symbol List, FNO Matket Data__ and __Market Holiday Data__ from excel itself in opened excel file.

<div align="center">
  <a href="https://github.com/kunalk3/NSE_Livedata_from_excel_extraction/issues"><img src="https://img.shields.io/github/issues/kunalk3/NSE_Livedata_from_excel_extraction" alt="Issues Badge"></a>
  <a href="https://github.com/kunalk3/NSE_Livedata_from_excel_extraction/graphs/contributors"><img src="https://img.shields.io/github/contributors/kunalk3/NSE_Livedata_from_excel_extraction?color=872EC4" alt="GitHub contributors"></a>
  <a href="https://www.python.org/downloads/release/python-390/"><img src="https://img.shields.io/static/v1?label=python&message=v3.9&color=faff00" alt="Python 3.9"</a>
  <a href="https://github.com/kunalk3/NSE_Livedata_from_excel_extraction/blob/main/LICENSE"><img src="https://img.shields.io/github/license/kunalk3/NSE_Livedata_from_excel_extraction?color=019CE0" alt="License Badge"/></a>
  <a href="https://github.com/kunalk3/NSE_Livedata_from_excel_extraction"><img src="https://img.shields.io/badge/lang-eng-ff1100"></img></a>
  <a href="https://github.com/kunalk3/NSE_Livedata_from_excel_extraction"><img src="https://img.shields.io/github/last-commit/kunalk3/NSE_Livedata_from_excel_extraction?color=309a02" alt="GitHub last commit">
</div>

<div align="center">   
  
  [![Windows](https://img.shields.io/badge/WindowsOS-000000?style=flat-square&logo=windows&logoColor=white)](https://www.microsoft.com/en-in/)
  [![Visual Studio Code](https://img.shields.io/badge/VSCode-0078d7.svg?style=flat-square&logo=visual-studio-code&logoColor=white)](https://code.visualstudio.com/)
  [![Jupyter](https://img.shields.io/badge/Jupyter-F37626.svg?style=flat-square&logo=Jupyter&logoColor=white)](https://jupyter.org/)
  [![Pycharm](https://img.shields.io/badge/Pycharm-41c907.svg?style=flat-square&logo=Pycharm&logoColor=white)](https://www.jetbrains.com/pycharm/)
  [![Colab](https://img.shields.io/badge/Colab-F9AB00.svg?style=flat-square&logo=googlecolab&logoColor=white)](https://colab.research.google.com/?utm_source=scs-index/)
  [![Spyder](https://img.shields.io/badge/Spyder-838485.svg?style=flat-square&logo=spyder%20ide&logoColor=white)](https://www.spyder-ide.org/)
</div>
  
---
## :books: Introduction NSE Market segments
- Pre-market -->
  
      ['NIFTY 50': 'NIFTY', 'NIFTY BANK': 'BANKNIFTY', 'EMERGE': 'SME', 'SECURITIES IN F&O': 'FO', 'OTHERS': 'OTHERS', 'ALL': 'ALL']
  
- Live Market -->
  
      'Broad Market Indices': ['NIFTY 50', 'NIFTY NEXT 50', 'NIFTY MIDCAP 50', 'NIFTY MIDCAP 100', 'NIFTY MIDCAP 150', 'NIFTY SMALLCAP 50', 'NIFTY SMALLCAP 100', 'NIFTY SMALLCAP 250', 
                              'NIFTY MIDSMALLCAP 400', 'NIFTY 100', 'NIFTY 200', 'NIFTY500 MULTICAP 50:25:25', 'NIFTY LARGEMIDCAP 250'],
      'Sectorial Indices': ['NIFTY AUTO','NIFTY BANK', 'NIFTY ENERGY', 'NIFTY FINANCIAL SERVICES', 'NIFTY FINANCIAL SERVICES 25/50', 'NIFTY FMCG', 'NIFTY IT', 'NIFTY MEDIA', 'NIFTY METAL', 
                           'NIFTY PHARMA', 'NIFTY PSU BANK', 'NIFTY REALTY', 'NIFTY PRIVATE BANK'], 
      'Others': ['Securities in F&O', 'Permitted to Trade'], 
      'Strategy Indices': ['NIFTY DIVIDEND OPPORTUNITIES 50', 'NIFTY50 VALUE 20', 'NIFTY100 QUALITY 30', 'NIFTY50 EQUAL WEIGHT', 'NIFTY100 EQUAL WEIGHT', 'NIFTY100 LOW VOLATILITY 30', 
                          'NIFTY ALPHA 50', 'NIFTY200 QUALITY 30', 'NIFTY ALPHA LOW-VOLATILITY 30', 'NIFTY200 MOMENTUM 30'],
      'Thematic Indices': ['NIFTY COMMODITIES', 'NIFTY INDIA CONSUMPTION', 'NIFTY CPSE', 'NIFTY INFRASTRUCTURE', 'NIFTY MNC', 'NIFTY GROWTH SECTORS 15', 'NIFTY PSE', 'NIFTY SERVICES SECTOR', 
                          'NIFTY100 LIQUID 15', 'NIFTY MIDCAP LIQUID 15']}

- Market Holidays -->
  
      ['Trading', 'Clearing']

---
  
## :bulb: Demo
- Below is the demonstrated sample at my local environments/ system with platform as __Visual Studio Code__ (V1.61.2) on OS __Windows 10__. 

https://user-images.githubusercontent.com/41562231/139540740-b667458d-8238-411b-a663-f6444a1cac7b.mp4

#### :pencil2: _Input_ - 
- Run python code __app.py__ and excel will open with naming conventition as _NSE_data_DDMMYYYY_. (Pre-requesite microsoft excel is installed on OS system)
- You are able to write _Keys_ as mentioned in excel first sheet. If Key name is matched, then only you are able to fetch the data, else no data is captured in real time from NSE website https://www.nseindia.com/ 
  
#### :bookmark: _Output_ - 
- `Data keys`
  
  ![MainScreen](https://user-images.githubusercontent.com/41562231/139541491-fa2cd7fc-322d-468e-bdb9-819de0b6af6a.JPG)
  
- `NSE Live-market data`
  
  ![LiveMarket](https://user-images.githubusercontent.com/41562231/139541452-77960989-8a86-4edc-955b-51474e202ce0.JPG)

- `NSE Pre-market data`
  
  ![PreMarket](https://user-images.githubusercontent.com/41562231/139541416-85dddccb-b015-4760-bdc8-ba9855735ddc.JPG)

- `NSE FnO market data`
  
  ![FnO](https://user-images.githubusercontent.com/41562231/139541350-75b5e377-5852-422a-bd7b-13174a83db3b.JPG)

---
  
## :wrench: Installation
- Create __virtual environment__ `python -m venv VIRTUAL_ENV_NAME` and activate it `.\VIRTUAL_ENV_NAME\Scripts\activate`.
- Install necessary library for this project from the file `requirements.txt` or manually install by `pip`.
  ```
  pip install -r requirements.txt
  ```
  To create project library requirements, use below command,
  ```
  pip freeze > requirements.txt
  ```
- In code, write your browser user-agent in headers before proceeding the script.
  ```python
  import requests
  
  def __init__(self):
        self.headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36'}
        self.session = requests.Session()
        self.session.get('http://nseindia.com', headers=self.headers)
  ```
  __How to capture the user-agent and API from NSE?__ _Refers below steps/ tutorial._
  
    1. Open NSE website https://www.nseindia.com/ and __inspect__ (Right click __->__ Inspect).
    2. Go to __Network__ and select www.nseindia.com under __Name__.
    3. Go to __Headers__ and copy the __user-agent__ value from bottom.

- Now run python file.
  ``` 
  python app.py
  ```
- Read more about __xlwings__ package and usages [__Excel automation__](https://www.xlwings.org/)

#### :computer: _py-to-exe conversion_ - 
- Need to install library [__pyinstaller__](https://pypi.org/project/pyinstaller/) `pip install pyinstaller`.
- Run below command (VS Code terminal) for _python file to exe conversion_. (Recommended: merge api.py code to app.py / make single python file)
  ```
  pyinstaller -i .\icon.ico -n NseAppication -w -F --log DEBUG .\app.py
  ``` 
    __-i__ icon , __-n__ application name , __-w__ windows based , __--log DEBUG__ logging if error occures , __app.py__ python file
- For __py to exe conversion__, refers [pyinstaller documentation](https://pyinstaller.readthedocs.io/en/stable/usage.html)
  
---  

## :bookmark: Directory Structure 
```bash
    .                                   # Root directory
    ├── app.py                          # Python code
    ├── api.py                          # Python code
    ├── icon.ico                        # Icon file
    ├── build                           # build folder generated due to py-to-exe conversion
    ├── dist                            # dist folder generated due to py-to-exe conversion
    │   ├── NSE_data_0102021.xlsx       # Outout excel working live file created after launching exe application 
    │   └── NSEApplication.exe          # Output exe application created after py-to-exe conversion (Not uploaded exe file due to application size  )
    ├── NseApplicaiton.spec             # Specification file generated due to py-to-exe conversion
    ├── requirements.txt                # Project requirements library with versions
    ├── README.md                       # Project README file
    └── LICENSE                         # Project LICENSE file
```

---  
  
## :iphone: Connect with me
`You say freak, I say unique. Don't wait for an opportunity, create it.`
  
__Let’s connect, share the ideas and feel free to ping me...__
  
<div align="center"> 
  <p align="left">
    <a href="https://linkedin.com/in/kunalkolhe3" target="blank"><img align="center" src="https://cdn.jsdelivr.net/npm/simple-icons@3.0.1/icons/linkedin.svg" alt="kunalkolhe3" height="30" width="40"/></a>
    <a href="https://github.com/kunalk3/" target="blank"><img align="center" src="https://cdn.jsdelivr.net/npm/simple-icons@3.0.1/icons/github.svg" alt="kunalkolhe3" height="30" width="40"/></a>
    <a href="mailto:kunalkolhe333@gmail.com" target="blank"><img align="center" src="https://cdn.jsdelivr.net/npm/simple-icons@3.0.1/icons/gmail.svg" alt="kunalkolhe333" height="30" width="40"/></a>
    <a href="https://www.hackerrank.com/kunalkolhe333" target="blank"><img align="center" src="https://cdn.jsdelivr.net/npm/simple-icons@3.0.1/icons/hackerrank.svg" alt="kunalkolhe333" height="30" width="40"/></a>
    <a href="https://fb.com/kunal.kolhe.98" target="blank"><img align="center" src="https://cdn.jsdelivr.net/npm/simple-icons@3.0.1/icons/facebook.svg" alt="kunal.kolhe.98" height="30" width="40"/></a>
    <a href="https://instagram.com/kkunalkkolhe" target="blank"><img align="center" src="https://cdn.jsdelivr.net/npm/simple-icons@3.0.1/icons/instagram.svg" alt="kkunalkkolhe" height="30" width="40"/></a>
  </p>
</div> 

