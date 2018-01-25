from webapp.models import *
from xlrd import open_workbook


def data():
    wb = open_workbook('E:\Chidhagni_python\django\myprog\Jan24,2018.xlsx')
    sheets=wb.nsheets

    ind=2
    while ind<sheets:
        print (ind)
        x = wb.sheet_by_index(ind)
        ind=ind+1
        print (ind)
        if(x.name=='NSELOW'):
            print("NSELOW")
            rows = x.nrows
            for i in range(1, rows):
                arr=x.row_values(i)[1:]
                db = NseLow(Symbol=arr[0], SecurityName=arr[1], New52L=arr[2], PreviousLow=arr[3],
                            PreviousLowDate=arr[4], LTP=arr[5], PreviousClose=arr[6], Change=arr[7],
                            PercentChange=arr[8], CurrentBusinessDate=arr[9])
                db.save()

        if(x.name=='BSE 52 HIGH'):
            print("BSE")
            rows = x.nrows
            for i in range(1,rows):
                arr=x.row_values(i)[1:]
                db=BseHigh(SecurtiyCode=arr[0],SecurityName=arr[1],LTP=arr[2],WeeksHigh52=arr[3],Previous52WeekHigh=arr[4],
                           AllTimeHigh=arr[5],CurrentBusinessDate=arr[6])
                db.save()

        if(x.name=='NSA CONTRACTS'):
            print("Contracts")
            rows=x.nrows
            for i in range(1,rows):
                arr=x.row_values(i)[1:]
                db=Contracts(Instrument =arr[0], Symbol=arr[1], Expiry=arr[2], StrikePrice=arr[3], Type=arr[4], LTP=arr[5], PrevClose=arr[6], PerChngLTP=arr[7],
                CBDOI=arr[8], PBDOI=arr[9], OIChange=arr[10], VolinContracts=arr[11], TurnOverinCr=arr[12], PremTrunOverInCr=arr[13], UnderlyningValue=arr[14],
                             TypeofOISpurts=arr[15], CurrentBusinessDate=arr[16], PreviousBusinessDate=arr[17])

                db.save()


        if(x.name=='OPTION CHAIN'):
            print("option chain")
            rows=x.nrows
            for i in range(1,rows):
                arr=x.row_values(i)[1:]
                db=OptionChain(CbdOI=arr[0],ChnginOI=arr[1],Volume=arr[2],IV=arr[3], LTP=arr[4], NetChng=arr[5],
                               BidQuant=arr[6], BidPrice=arr[7],AskPrice = arr[8],AskQuant = arr[9],StrikePrice=arr[10],
                               CurrentBusinessDate=arr[11],Category=arr[12],ExpiryDate=arr[13],StockMarketIndex=arr[14])

                db.save()




