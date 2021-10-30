from api import NSEIndia, NSEIndia2
import xlwings as xw
import os, time, sys
import pandas as pd
import tkinter as tk
from datetime import date


nse = NSEIndia()
nse2 = NSEIndia2()

# GUI section
root = tk.Tk()
frame = tk.Frame(root)
frame.pack()
root.title('NSE live data')
root.iconbitmap('./icon.ico')
root.geometry('250x100')
root.minsize(150, 100)
root.maxsize(300, 110)


cwd = os.getcwd()
today = date.today()
current_date = today.strftime("%d%m%Y")
print(current_date)
cwdFile = 'NSE_data_' + current_date + '.xlsx'
location = os.path.join(cwd, cwdFile)
print(cwdFile)
print(os.path.join(cwd, cwdFile))


if not os.path.exists(cwdFile):
    try:
        wb = xw.Book()
        wb.sheets.add('SymbolListData')
        wb.sheets.add('HolidayData')
        wb.sheets.add('OptionChainData')
        wb.sheets.add('LiveMarketData')
        wb.sheets.add('PreMarketData')
        wb.sheets.add('DataKeys')
        wb.save(cwdFile)
        wb.close()
    except Exception as err:
        print(f'Error while creating NSE data excel : {err}')
else:
    pass    

def onClickStartButton():
    time.sleep(2)

    wb = xw.Book(cwdFile)
    app = xw.apps.active  
    dk = wb.sheets('DataKeys')
    pdd = wb.sheets('PreMarketData')
    ld = wb.sheets('LiveMarketData')
    oc = wb.sheets('OptionChainData')
    hd = wb.sheets('HolidayData')
    sld = wb.sheets('SymbolListData')

    # First excel page reference (DataKeys)
    dk.range('a:b').value = None 
    dk.range('d:e').value = None
    dk.range('g:h').value = None
    dk.range('j:k').value = None
    dk.range('m:n').value = None
    dk.range('a1').value = pd.DataFrame({'Pre-Market Keys': nse.pre_market_keys})
    dk.range('d1').value = pd.DataFrame({'Live Market Keys': nse.live_market_keys})
    dk.range('g1').value = pd.DataFrame({'Holiday Keys': nse.holiday_keys})
    dk.range('j1').value = pd.DataFrame({'FNO Symbols': ['NIFTY', 'BANKNIFTY'] + 
                                        nse.NseLiveMarketData("Securities in F&O", symbol_list=True)})
    dk.range('m1').value = 'Exit here (Q/q)'
    quite_input = dk.range("m1")
    quite_input.autofit()

    # Pre-market key heading
    pdd.range('d2').value = 'Pre-Market Key:-'
    pm_input_section = pdd.range("d2")
    pm_input_section.autofit()
    pm_input_section.color = (0,0,0)
    pm_input_section.api.Font.Color = 0xFFFFFF
    pm_input_section.api.Font.Bold = True
    pm_input_section.api.Font.Size = 11
    pm_input_section.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    pm_input_section.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
    pdd.range('a4:z2000').value = None
    previous_pre_key = None

    # Live market key heading
    ld.range('d2').value = 'Live Market Key:-'
    ld_input_section = ld.range("d2")
    ld_input_section.autofit()
    ld_input_section.color = (0,0,0)
    ld_input_section.api.Font.Color = 0xFFFFFF
    ld_input_section.api.Font.Bold = True
    ld_input_section.api.Font.Size = 11
    ld_input_section.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    ld_input_section.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
    ld.range('a4:z2000').value = None
    previous_live_key = None

    # Option chain symbol key heading
    oc.range('d2').value = 'Symbol key:-'
    oc_input_section = oc.range("d2")
    oc_input_section.autofit()
    oc_input_section.color = (0,0,0)
    oc_input_section.api.Font.Color = 0xFFFFFF
    oc_input_section.api.Font.Bold = True
    oc_input_section.api.Font.Size = 11
    oc_input_section.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    oc_input_section.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
    oc.range('a4:z2000').value = None
    previous_oc_symbol = None

    # Symbols key heading
    sld.range('d2').value = 'Symbol Key:-'
    sld_input_section = sld.range("d2")
    sld_input_section.autofit()
    sld_input_section.color = (0,0,0)
    sld_input_section.api.Font.Color = 0xFFFFFF
    sld_input_section.api.Font.Bold = True
    sld_input_section.api.Font.Size = 11
    sld_input_section.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    sld_input_section.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
    sld.range('a4:d2000').value = None
    pre_symbol_list = None

    # Holiday key heading
    hd.range('d2').value = 'Holiday Key:-'
    hd_input_section = hd.range("d2")
    hd_input_section.autofit()
    hd_input_section.color = (0,0,0)
    hd_input_section.api.Font.Color = 0xFFFFFF
    hd_input_section.api.Font.Bold = True
    hd_input_section.api.Font.Size = 11
    hd_input_section.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    hd_input_section.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
    hd.range('a4:e500').value = None
    pre_holiday_keys = None

    while True:
        time.sleep(1)

        ## ----------- NSE holiday section ----------- ##
        try:
            holiday_key = hd.range('e2').value
            hd.autofit()
            hd.range('e2').color = (194,194,194)
            holiday_key = holiday_key.lower()
                
            if pre_holiday_keys != holiday_key:
                hd.range('a4:z2000').value = None
                pre_holiday_keys = holiday_key
                if holiday_key is not None:
                    hd.range('a4').value = nse.NseHoliday(holiday_key)
                    hd.range('g2').value = ''

                    # Sheet formating
                    hd_header = hd.range("a4").expand('right')
                    hd_header.color = (112,173,71)
                    hd_header.api.Font.Color = 0xFFFFFF
                    hd_header.api.Font.Bold = True
                    # hd_header.api.Font.Name = 'Arial'
                    hd_header.api.Font.Size = 11
                    hd_header.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                    hd_header.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
                    hd_header.api.WrapText = True

                    hd_col_dal = hd.range("a5").expand('table')
                    hd_col_dal.color = (194,194,194)
                    for border_id in range(7,13):
                        hd_col_dal.api.Borders(border_id).Weight = 1
                        hd_col_dal.api.Borders(border_id).Color = 0xFFFFFF
            else:
                hd.range('g2').value = 'N.A'   
                hd.range('4:2000').api.Delete()  # hd.range('a4:z2000').value = None # (Deletion instead of None)
        except:
            pass
        
        ## ----------- NSE live market section ----------- ##
        try:
            symbol_keys = sld.range('e2').value
            sld.autofit()
            sld.range('e2').color = (194,194,194)
            symbol_keys = symbol_keys.upper()
            if pre_symbol_list != symbol_keys:
                sld.range('a4:b2000').value = None
                pre_symbol_list = symbol_keys
                if pre_symbol_list is not None:
                    sld.range('a4').value = pd.DataFrame({'Symbols Names': nse.NseLiveMarketData(symbol_keys, symbol_list=True)}) 
                    sld.range('g2').value = None
            else:
                sld.range('g2').value = 'N.A'
                sld.range('4:2000').api.Delete()
        except:
            pass

        ## ----------- NSE option chain section ----------- ##
        try:
            oc_symbol = oc.range('e2').value
            oc.autofit()
            oc.range('e2').color = (194,194,194)
            oc_symbol = oc_symbol.upper()
            if previous_oc_symbol is None:
                previous_oc_symbol = oc_symbol
            if previous_oc_symbol != oc_symbol:
                oc.range('a4:z2000').value = None
                previous_oc_symbol = oc_symbol
            if oc_symbol is not None:
                if oc_symbol == 'NIFTY' or oc_symbol == 'BANKNIFTY':
                    indices = True
                    oc.range('g2').value = None
                else:
                    indices = False
                oc.range('a4').value = nse2.GetOptionChainData(oc_symbol, indices)
                oc.range('g2').value = ''
                oc.autofit()

                oc_header = oc.range("a4").expand('right')
                oc_header.color = (112,173,71)
                oc_header.api.Font.Color = 0xFFFFFF
                oc_header.api.Font.Bold = True
                oc_header.api.Font.Size = 11
                oc_header.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                oc_header.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
                oc_header.api.WrapText = True

                oc_col_dal = oc.range("a5").expand('table')
                oc_col_dal.color = (194,194,194)
                for border_id in range(7,13):
                    oc_col_dal.api.Borders(border_id).Weight = 1
                    oc_col_dal.api.Borders(border_id).Color = 0xFFFFFF
            else:
                oc.range('g2').value = 'N.A'
                oc.range('4:2000').api.Delete()
        except:
            pass
        
        ## ----------- NSE pre market section ----------- ##
        try:
            global pre_market_key
            pre_market_key = pdd.range('e2').value
            pdd.autofit()
            pdd.range('e2').color = (194,194,194)
            pre_market_key = pre_market_key.upper()
            if previous_pre_key != pre_market_key:
                pdd.range('a4:z2000').value = None
                previous_live_key = pre_market_key
                
                if pre_market_key is not None:
                    pdd.range('a4').value = nse.NsePreMarketData(pre_market_key)
                    pdd.range('g2').value = ''
                    pdd.autofit()

                    pdd_header = pdd.range("a4").expand('right')
                    pdd_header.color = (112,173,71)
                    pdd_header.api.Font.Color = 0xFFFFFF
                    pdd_header.api.Font.Bold = True
                    pdd_header.api.Font.Size = 11
                    pdd_header.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                    pdd_header.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
                    pdd_header.api.WrapText = True

                    pdd_col_dal = pdd.range("a5").expand('table')
                    pdd_col_dal.color = (194,194,194)
                    for border_id in range(7,13):
                        pdd_col_dal.api.Borders(border_id).Weight = 1
                        pdd_col_dal.api.Borders(border_id).Color = 0xFFFFFF
            else:
                pdd.range('g2').value = 'N.A'
                pdd.range('3:4000').api.Delete()
        except:
           pass

        ## ----------- NSE symbol list section ----------- ##
        try:
            live_market_key = ld.range('e2').value
            ld.autofit()
            ld.range('e2').color = (194,194,194)
            live_market_key = live_market_key.upper()
            if previous_live_key is None:
                previous_live_key = live_market_key
                ld.range('g2').value = 'N.A'
            if previous_live_key != live_market_key:
                ld.range('a4:z2000').value = None
                previous_live_key = live_market_key
            if live_market_key is not None:
                ld.range('a4').value = nse.NseLiveMarketData(live_market_key, symbol_list=False)
                ld.range('g2').value = ''
                ld.autofit()

                ld_header = ld.range("a4").expand('right')
                ld_header.color = (112,173,71)
                ld_header.api.Font.Color = 0xFFFFFF
                ld_header.api.Font.Bold = True
                ld_header.api.Font.Size = 11
                ld_header.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                ld_header.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
                ld_header.api.WrapText = True

                ld_col_dal = ld.range("a5").expand('table')
                ld_col_dal.color = (194,194,194)
                for border_id in range(7,13):
                    ld_col_dal.api.Borders(border_id).Weight = 1
                    ld_col_dal.api.Borders(border_id).Color = 0xFFFFFF
            else:
                ld.range('g2').value = 'N.A'
                ld.range('4:4000').api.Delete()    
        except:
            pass

        ## ----------- Quite from excel sestion ----------- ##
        try:
            quite = dk.range('m2').value
            dk.range('m2').color = (194,194,194)
            quite = quite.lower()
            print('val', quite)
            if quite == 'Q' or quite == 'q':
                app.quit()
                break 
        except:
            pass   

def onExit():
    sys.exit(0)

quiteButton = tk.Button(frame, text="QUIT", fg="red", command=onExit)
quiteButton.pack(side=tk.LEFT, padx=15, pady=20)

startButton = tk.Button(frame, text="START", command=onClickStartButton)
startButton.pack(side=tk.LEFT, padx=15, pady=20)


root.mainloop()



