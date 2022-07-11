# -*- coding: utf-8 -*
import os
import time
import sys
from threading import Thread

import comtypes
import comtypes.client
from comtypes.client import PumpEvents

client = comtypes.client.CreateObject("COM_PFCFAPI.COM_PFCFAPI")

ACTNO = ""
PRODUCTID = ""


class TimeInForceEnum(object):
    FOK = 70
    IOC = 73
    ROD = 82


class OpenCloseEnum(object):
    Y = 48
    N = 49


class DayTradeEnum(object):
    Y = 89
    N = 78


class ReplaceTypeEnum(object):
    Decrease = 0
    Cancel = 1
    ChangePrice = 2


class OrderTypeEnum(object):
    Stop = 51
    StopLimit = 52
    Limit = 76
    Market = 77


class SideEnum(object):
    Buy = 66
    Sell = 83


class CPEnum(object):
    NONE = 32
    Call = 67
    Future = 70
    Put = 80


class EventSink(object):
    """ PFC events """

    def ICOM_PFCFAPI_Events_PFCloginStatus(self, msg):
        print("登入狀態 OnPFCloginStatus: %s\n" % msg)
        return True

    def ICOM_PFCFAPI_Events_PFCErrorData(self, msg):
        print("OnPFCErrorData: %s" % msg)
        return True

    def ICOM_PFCFAPI_Events_PFCFutures(self, Class, COMMODITYID, desc, month, MaxPrice, MinPrice):
        print("OnPFCFutures")
        return True

    def ICOM_PFCFAPI_Events_PFCOptions(self, Class, COMMODITYID, desc, month, CP, STRIKEPRICE, MaxPrice, MinPrice):
        print("OnPFCOptions")
        return True

    """ ACCOUNT events """

    def ICOM_PFCFAPI_Events_FAccountOnDisconnected(self):
        print("OnFAccountOnDisconnected")
        return True

    def ICOM_PFCFAPI_Events_FAccountOnConnected(self):
        print("已連上外期帳務主機事件 OnFAccountOnConnected")
        return True

    def ICOM_PFCFAPI_Events_FAccountOnMarginError(self, ERRORCODE, MESSAGE):
        print("OnFAccountOnMarginError")
        return True

    def ICOM_PFCFAPI_Events_FAccountOnPositionError(self, ERRORCODE, MESSAGE):
        print("OnFAccountOnPositionError")
        return True

    def ICOM_PFCFAPI_Events_FAccountOnUnLiquidationDetailError(self, ERRORCODE, MESSAGE):
        print("OnFAccountOnUnLiquidationDetailError")
        return True

    def ICOM_PFCFAPI_Events_FAccountOnMarginData(self, marginData):
        return True

    def ICOM_PFCFAPI_Events_FAccountOnPositionData(self, positionData):
        return True

    def ICOM_PFCFAPI_Events_FAccountOnUnLiquidationDetail(self, unLiquidationDetail):
        return True

    def ICOM_PFCFAPI_Events_FAccountOnMarginDataPython(self, TCNT, NCNT, WEBID, NMVSEQ, CURRENCY, LCTDAB, ORIGNFEE,
                                                       TAXAMT, CTAXAMT, DWAMT, OSPRTLOS, PRTLOS, BMKTVAL, SMKTVAL,
                                                       OPREMIUM, TPREMIUM, EQUITY, IAMT, MAMT, EXCESS, ORDEXCESS,
                                                       TRUSTPRICE, EXCERCISEPRICE, DRIGHTPRICE, Time, OVAMT, YSDPRTLOS,
                                                       PRTLOSAMT, MKAMT, STRIKEAMT, NCTDAB, ACCOUNTAMT, ONAMT,
                                                       TOTALRISK, RISK, KEEPRATE):
        print("收到外期保證金查詢結果: %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
              "%s,%s,%s,%s,%s,%s,%s,%s,%s" % (TCNT, NCNT, WEBID, NMVSEQ, CURRENCY, LCTDAB, ORIGNFEE, TAXAMT, CTAXAMT,
                                              DWAMT, OSPRTLOS, PRTLOS, BMKTVAL, SMKTVAL, OPREMIUM, TPREMIUM, EQUITY,
                                              IAMT, MAMT, EXCESS, ORDEXCESS, TRUSTPRICE, EXCERCISEPRICE, DRIGHTPRICE,
                                              Time, OVAMT, YSDPRTLOS,PRTLOSAMT, MKAMT, STRIKEAMT, NCTDAB, ACCOUNTAMT,
                                              ONAMT, TOTALRISK, RISK, KEEPRATE))
        return True

    def ICOM_PFCFAPI_Events_FAccountOnPositionDataPython(self, TCNT, NCNT, WEBID, NMVSEQ, FIRM, ACTNO, EXH, COMTYPE,
                                                         COMNO, COMYM, STKPRC, CALLPUT, BIQTY, SIQTY, BOTQTY, SOTQTY,
                                                         BMQTY, SMQTY, BOQTY, SOQTY, INCOMEDATE, CUR, MATCHPRICE,
                                                         REALPRICE, PRTLOS, BALANCEQTY, INCOMEBALANCE):
        print("收到外期即時部位查詢結果: %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s"
              % (TCNT, NCNT, WEBID, NMVSEQ, FIRM, ACTNO, EXH, COMTYPE,
                                                         COMNO, COMYM, STKPRC, CALLPUT, BIQTY, SIQTY, BOTQTY, SOTQTY,
                                                         BMQTY, SMQTY, BOQTY, SOQTY, INCOMEDATE, CUR, MATCHPRICE,
                                                         REALPRICE, PRTLOS, BALANCEQTY, INCOMEBALANCE))
        return True

    def ICOM_PFCFAPI_Events_FAccountOnUnLiquidationDetailPython(self, TCNT, NCNT, WEBID, NMVSEQ, FIRM, ACTNO, EXH,
                                                                SEQNO, ORDNO1, TRDNO1, DIVIDESEQ1, TRDDT1, PS1,
                                                                COMTYPE1, COMNO1, COMYM1, STKPRC1, CALLPUT1, OTQTY1,
                                                                TRDPRE1, RELPRE1, PRTLOS1, IAMT1, MAMT1, SPREAD,
                                                                CURRENCY1, TRDPRC1, ORDNO2, TRDNO2, DIVIDESEQ2, TRDDT2,
                                                                PS2, COMTYPE2, COMNO2, COMYM2, STKPRC2, CALLPUT2,
                                                                OTQTY2, TRDPRE2, RELPRE2, PRTLOS2, IAMT2, MAMT2,
                                                                CURRENCY2, TRDPRC2, FCMNO, DTRADE, TRDTYPE, COMBOQTY1,
                                                                COMBOQTY2, ORDBROKER, SETTDATE, CLOSEDATE, ORDTYPE,
                                                                NETWORKID, PRTLOSTW1, PRTLOSTW2, FEES1, FEES2, BUSTAX1,
                                                                BUSTAX2, NPRTLOS1, NPRTLOS2, NPRTLOSNT1, NPRTLOSNT2):
        print("收到外期未平倉明細查詢結果: %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
              "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
              "%s,%s,%s,%s,%s" % (TCNT, NCNT, WEBID, NMVSEQ, FIRM, ACTNO, EXH,
                                                                SEQNO, ORDNO1, TRDNO1, DIVIDESEQ1, TRDDT1, PS1,
                                                                COMTYPE1, COMNO1, COMYM1, STKPRC1, CALLPUT1, OTQTY1,
                                                                TRDPRE1, RELPRE1, PRTLOS1, IAMT1, MAMT1, SPREAD,
                                                                CURRENCY1, TRDPRC1, ORDNO2, TRDNO2, DIVIDESEQ2, TRDDT2,
                                                                PS2, COMTYPE2, COMNO2, COMYM2, STKPRC2, CALLPUT2,
                                                                OTQTY2, TRDPRE2, RELPRE2, PRTLOS2, IAMT2, MAMT2,
                                                                CURRENCY2, TRDPRC2, FCMNO, DTRADE, TRDTYPE, COMBOQTY1,
                                                                COMBOQTY2, ORDBROKER, SETTDATE, CLOSEDATE, ORDTYPE,
                                                                NETWORKID, PRTLOSTW1, PRTLOSTW2, FEES1, FEES2, BUSTAX1,
                                                                BUSTAX2, NPRTLOS1, NPRTLOS2, NPRTLOSNT1, NPRTLOSNT2))
        return True

    """ QUOTE events """

    def ICOM_PFCFAPI_Events_FQuoteOnDisconnected(self):
        print("OnFQuoteOnDisconnected")
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnConnected(self):
        print("已連上外期行情主機事件 OnFQuoteOnConnected")
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataTrade(self, val):
        # print("OnFQuoteOnTickDataTrade")
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataTradePython(self, EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                        DISPLAYMULTIPLY, Total, LastPrice, LastVolume, Time):
        print("收到外期成交價:%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s" % (EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                        DISPLAYMULTIPLY, Total, LastPrice, LastVolume, Time))
        FQuoteUnRegItem()
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataBid(self, val):
        # print("OnFQuoteOnTickDataBid")
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataBidPython(self, EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                      DISPLAYMULTIPLY, BidDOM1Price, BidDOM1Volume, BidDOM2Price,
                                                      BidDOM2Volume, BidDOM3Price, BidDOM3Volume, BidDOM4Price,
                                                      BidDOM4Volume, BidDOM5Price, BidDOM5Volume, BidDOM6Price,
                                                      BidDOM6Volume, BidDOM7Price, BidDOM7Volume, BidDOM8Price,
                                                      BidDOM8Volume, BidDOM9Price, BidDOM9Volume, BidDOM10Price,
                                                      BidDOM10Volume):
        print("收到外期最佳買價%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s"
              % (EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                      DISPLAYMULTIPLY, BidDOM1Price, BidDOM1Volume, BidDOM2Price,
                                                      BidDOM2Volume, BidDOM3Price, BidDOM3Volume, BidDOM4Price,
                                                      BidDOM4Volume, BidDOM5Price, BidDOM5Volume, BidDOM6Price,
                                                      BidDOM6Volume, BidDOM7Price, BidDOM7Volume, BidDOM8Price,
                                                      BidDOM8Volume, BidDOM9Price, BidDOM9Volume, BidDOM10Price,
                                                      BidDOM10Volume))
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataOfferPython(self, val, EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                        DISPLAYMULTIPLY, Total, LastPrice, LastVolume, Time):
        print("收到外期成交價:%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s" % (EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                        DISPLAYMULTIPLY, Total, LastPrice, LastVolume, Time))
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataOffer(self, val, ):
        # print("OnFQuoteOnTickDataOffer")
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataImpliedPython(self, EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                          DISPLAYMULTIPLY, ImpliedBidPrice, ImpliedBidVolume,
                                                          ImpliedOfferPrice, ImpliedOfferVolume):
        print("收到外期隱含買賣價量:%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s" % (EXCHAGE, SYMBOL, YM, CP, STRIKE,
                                                               DISPLAYDENOMINATOR, DISPLAYMULTIPLY, ImpliedBidPrice,
                                                               ImpliedBidVolume, ImpliedOfferPrice, ImpliedOfferVolume))
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataImplied(self, val):
        # print("OnFQuoteOnTickDataImplied")
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataHighLowPython(self, EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                          DISPLAYMULTIPLY, High, Low):
        print("收到外期最高最低價%s,%s,%s,%s,%s,%s,%s,%s,%s" % (EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                        DISPLAYMULTIPLY, High, Low))
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataHighLow(self, val):
        # print("OnFQuoteOnTickDataHighLow")
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataOpenclosePython(self, EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                            DISPLAYMULTIPLY, Opening, Closing):
        print("收到外期開收盤價%s,%s,%s,%s,%s,%s,%s,%s,%s" % (EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAYDENOMINATOR,
                                                      DISPLAYMULTIPLY, Opening, Closing))
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataOpenclose(self, val):
        # print("OnFQuoteOnTickDataOpenclose")
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataSettlePython(self, EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAY_DENOMINATOR,
                                                         DISPLAY_MULTIPLY, CurrStl, NewStl):
        print("收到外期結算價%s,%s,%s,%s,%s,%s,%s,%s,%s" % (EXCHAGE, SYMBOL, YM, CP, STRIKE, DISPLAY_DENOMINATOR,
                                                         DISPLAY_MULTIPLY, CurrStl, NewStl))
        return True

    def ICOM_PFCFAPI_Events_FQuoteOnTickDataSettle(self, val):
        # print("OnFQuoteOnTickDataSettle")
        return True

    """ TRADE events """

    def ICOM_PFCFAPI_Events_FTradeOnQueryReply(self, count, recordno, freplydata):
        return True

    def ICOM_PFCFAPI_Events_FTradeOnQueryMatch(self, count, recordno, fmatchdata):
        return True

    def ICOM_PFCFAPI_Events_FTradeOnReply(self, freplydata):
        return True

    def ICOM_PFCFAPI_Events_FTradeOnMatch(self, fmatchdata):
        return True

    def ICOM_PFCFAPI_Events_FTradeOnQueryReplyPython(self, count, recordno, TRADEDATE, ORDERNO, ORDERTIME, BROKERID,
                                                     INVESTORACNO, SUBACT, BS, PRODUCTKIND, EXCHANGE, SYMBOL1,
                                                     MATURITYMONTHYEAR1, PUTORCALL1, STRIKEPRICE1, SIDE1, SYMBOL2,
                                                     MATURITYMONTHYEAR2, PUTORCALL2, STRIKEPRICE2, SIDE2, PRICE,
                                                     STOPPRICE, ORDERQTY, MATCHQTY, NOMATCHQTY, DELQTY, STATUSCODE,
                                                     ORDERSTATUS, NETWORKID, TIMEINFORCE, OPENCLOSE, ORDERTYPE,
                                                     EXPIREDATE, DTRADE, MDATE, SEQ, SOURCECODE, NOTE, LASTEXECID,
                                                     LASTORDEXECID, LASTMATCHPRICE, LASTMATCHQTY):
        print("收到查詢回報 %s %s %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
              "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s" % (count, recordno, TRADEDATE, ORDERNO, ORDERTIME, BROKERID,
                                                     INVESTORACNO, SUBACT, BS, PRODUCTKIND, EXCHANGE, SYMBOL1,
                                                     MATURITYMONTHYEAR1, PUTORCALL1, STRIKEPRICE1, SIDE1, SYMBOL2,
                                                     MATURITYMONTHYEAR2, PUTORCALL2, STRIKEPRICE2, SIDE2, PRICE,
                                                     STOPPRICE, ORDERQTY, MATCHQTY, NOMATCHQTY, DELQTY, STATUSCODE,
                                                     ORDERSTATUS, NETWORKID, TIMEINFORCE, OPENCLOSE, ORDERTYPE,
                                                     EXPIREDATE, DTRADE, MDATE, SEQ, SOURCECODE, NOTE, LASTEXECID,
                                                     LASTORDEXECID, LASTMATCHPRICE, LASTMATCHQTY))
        return True

    def ICOM_PFCFAPI_Events_FTradeOnQueryMatchPython(self, count, recordno, TRADEDATE, ORDERNO, MATCHTIME, BROKERID,
                                                     INVESTORACNO, SUBACT, BS, PRODUCTKIND, EXCHANGE, SYMBOL1,
                                                     MATURITYMONTHYEAR1, PUTORCALL1, STRIKEPRICE1, SIDE1, MATCHPRICE,
                                                     MATCHQTY, SYMBOL2, MATURITYMONTHYEAR2, PUTORCALL2, STRIKEPRICE2,
                                                     SIDE2, PRICE1, PRICE2, EXECID, ORDEXECID, NETWORKID, OPENCLOSE,
                                                     NOTE):
        print("收到查詢成回 %s %s %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s"
              % (count, recordno, TRADEDATE, ORDERNO, MATCHTIME, BROKERID,
                                                     INVESTORACNO, SUBACT, BS, PRODUCTKIND, EXCHANGE, SYMBOL1,
                                                     MATURITYMONTHYEAR1, PUTORCALL1, STRIKEPRICE1, SIDE1, MATCHPRICE,
                                                     MATCHQTY, SYMBOL2, MATURITYMONTHYEAR2, PUTORCALL2, STRIKEPRICE2,
                                                     SIDE2, PRICE1, PRICE2, EXECID, ORDEXECID, NETWORKID, OPENCLOSE,
                                                     NOTE))
        return True

    def ICOM_PFCFAPI_Events_FTradeOnReplyPython(self, TRADEDATE, ORDERNO, ORDERTIME, BROKERID, INVESTORACNO, SUBACT, BS,
                                                PRODUCTKIND, EXCHANGE, SYMBOL1, MATURITYMONTHYEAR1, PUTORCALL1,
                                                STRIKEPRICE1, SIDE1, SYMBOL2, MATURITYMONTHYEAR2, PUTORCALL2,
                                                STRIKEPRICE2, SIDE2, PRICE, STOPPRICE, ORDERQTY, MATCHQTY, NOMATCHQTY,
                                                DELQTY, STATUSCODE, ORDERSTATUS, NETWORKID, TIMEINFORCE, OPENCLOSE,
                                                ORDERTYPE, EXPIREDATE, DTRADE, MDATE, SEQ, SOURCECODE, NOTE, LASTEXECID,
                                                LASTORDEXECID, LASTMATCHPRICE, LASTMATCHQTY):
        print("即時回報  委託書號%s 資料 %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
              "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s" % (ORDERNO, TRADEDATE, ORDERTIME, BROKERID, INVESTORACNO, SUBACT, BS,
                                                PRODUCTKIND, EXCHANGE, SYMBOL1, MATURITYMONTHYEAR1, PUTORCALL1,
                                                STRIKEPRICE1, SIDE1, SYMBOL2, MATURITYMONTHYEAR2, PUTORCALL2,
                                                STRIKEPRICE2, SIDE2, PRICE, STOPPRICE, ORDERQTY, MATCHQTY, NOMATCHQTY,
                                                DELQTY, STATUSCODE, ORDERSTATUS, NETWORKID, TIMEINFORCE, OPENCLOSE,
                                                ORDERTYPE, EXPIREDATE, DTRADE, MDATE, SEQ, SOURCECODE, NOTE, LASTEXECID,
                                                LASTORDEXECID, LASTMATCHPRICE, LASTMATCHQTY))
        return True

    def ICOM_PFCFAPI_Events_FTradeOnMatchPython(self, TRADEDATE, ORDERNO, MATCHTIME, BROKERID, INVESTORACNO, SUBACT, BS,
                                                PRODUCTKIND, EXCHANGE, SYMBOL1, MATURITYMONTHYEAR1, PUTORCALL1,
                                                STRIKEPRICE1, SIDE1, MATCHPRICE, MATCHQTY, SYMBOL2, MATURITYMONTHYEAR2,
                                                PUTORCALL2, STRIKEPRICE2, SIDE2, PRICE1, PRICE2, EXECID, ORDEXECID,
                                                NETWORKID, OPENCLOSE, NOTE):
        print("即時成回 %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s"
              % (TRADEDATE, ORDERNO, MATCHTIME, BROKERID, INVESTORACNO, SUBACT, BS,
                                                PRODUCTKIND, EXCHANGE, SYMBOL1, MATURITYMONTHYEAR1, PUTORCALL1,
                                                STRIKEPRICE1, SIDE1, MATCHPRICE, MATCHQTY, SYMBOL2, MATURITYMONTHYEAR2,
                                                PUTORCALL2, STRIKEPRICE2, SIDE2, PRICE1, PRICE2, EXECID, ORDEXECID,
                                                NETWORKID, OPENCLOSE, NOTE))
        return True

    def ICOM_PFCFAPI_Events_FTradeOnDisconnected(self):
        print("OnFTradeOnDisconnected")
        return True

    def ICOM_PFCFAPI_Events_FTradeOnConnected(self):
        print("已連上外期交易主機事件 OnFTradeOnConnected")
        return True

    """ ACCOUNT events2 """

    def ICOM_PFCFAPI_Events_DAccountOnDisconnected(self):
        print("OnDQuoteOnDisconnected")
        return True

    def ICOM_PFCFAPI_Events_DAccountOnConnected(self):
        print("已連上內期帳務主機事件 OnDAccountOnConnected")
        return True

    def ICOM_PFCFAPI_Events_DAccountOnMarginData(self, EXRATE, LCTDAB, LTDAB, DWAMT, OSPRTLOS, PRTLOS, OPTOSPRTLOS,
                                                 OPTPRTLOS, TPREMIUM,
                                                 ORIGNFEE, CTAXAMT, ORDPREMIUM, CTDAB, ORDIAMT, IAMT, MAMT, ORDCEXCESS,
                                                 BPREMIUM,
                                                 SPREMIUM, OPTEQUITY, INIRATE, MATRATE, OPTRATE, TWDOPTEQUITY,
                                                 TWDINIRATE, TWDORDEXCESS,
                                                 TMP1PRICES, EXCERCISEPRICE, SYSDATE, SYSTIME):
        print("匯率 %s"
              "\t昨日權益數 %s"
              "\t昨日餘額 %s"
              "\t存提 %s"
              "\t本日期貨平倉損益淨額 %s"
              "\t未沖銷期貨浮動損益 %s"
              "\t選擇權平倉損益 %s"
              "\t選擇權未平倉浮動損益 %s"
              "\t權利金收入與支出 %s"
              "\t手續費 %s"
              "\t期交稅 %s"
              "\t委託權利金 %s"
              "\t權益數 %s"
              "\t委託保證金 %s"
              "\t原始保證金 %s"
              "\t維持保證金 %s"
              "\t可動用（出金）保證金 %s"
              "\t 未沖銷買方權利金市值 %s"
              "\t未沖銷賣方權利金市值 %s"
              "\t權益總值 %s"
              "\t原始比率 %s"
              "\t維持比率 %s"
              "\t風險指標 %s"
              "\t台幣權益總值 %s"
              "\t台幣原始比率 %s"
              "\t台幣下單可用保證金 %s"
              "\t依「加收保證金指標」所加收之保證金 %s"
              "\t到期履約損益 %s"
              "\t資料更新日期 %s"
              "\t資料更新時間 %s" % (EXRATE, LCTDAB, LTDAB, DWAMT, OSPRTLOS, PRTLOS, OPTOSPRTLOS, OPTPRTLOS,
                                      TPREMIUM, ORIGNFEE, CTAXAMT, ORDPREMIUM, CTDAB, ORDIAMT, IAMT, MAMT, ORDCEXCESS,
                                      BPREMIUM, SPREMIUM, OPTEQUITY, INIRATE, MATRATE, OPTRATE, TWDOPTEQUITY,
                                      TWDINIRATE, TWDORDEXCESS, TMP1PRICES, EXCERCISEPRICE, SYSDATE, SYSTIME))
        return True

    def ICOM_PFCFAPI_Events_DAccountOnMarginError(self, ERRORCODE, MESSAGE):
        print("OnDAccountOnMarginError %s %s" % (ERRORCODE, MESSAGE))
        return True

    def ICOM_PFCFAPI_Events_DAccountOnPositionData(self, count, recordno, investorAcno, ProductId, productKind, OTQtyB,
                                                   OTQtyS,
                                                   NowOrderQtyB, NowOrderQtyS, NowMatchQtyB, NowMatchQtyS, TodayEnd,
                                                   NowOTQtyB,
                                                   NowOTQtyS, RealPrice, AvgCostB, AvgCostS, PriceDiffB, PriceDiffS,
                                                   PricePL, Curren,
                                                   LiquidationPL):
        print("筆數 %s"
              "\t目前第幾筆 %s"
              "\t帳號 %s"
              "\t商品代碼 %s"
              "\t商品種類 %s"
              "\t昨日買進留倉 %s"
              "\t昨日賣出留倉 %s"
              "\t今日委託買進 %s"
              "\t今日委託賣出 %s"
              "\t今日成交買進 %s"
              "\t今日成交賣出 %s"
              "\t本日了結 %s"
              "\t目前買進留倉 %s"
              "\t目前賣出留倉 %s"
              "\t參考即時價 %s"
              "\t買進平均成交價 %s"
              "\t賣出平均成交價 %s"
              "\t價差買 %s"
              "\t價差賣 %s"
              "\t未平倉損益 %s"
              "\t幣別 %s"
              "\t平倉損益 %s" % (count, recordno, investorAcno, ProductId, productKind, OTQtyB, OTQtyS,
                                          NowOrderQtyB, NowOrderQtyS, NowMatchQtyB, NowMatchQtyS, TodayEnd, NowOTQtyB,
                                          NowOTQtyS, RealPrice, AvgCostB, AvgCostS, PriceDiffB, PriceDiffS, PricePL,
                                          Curren, LiquidationPL))
        return True

    def ICOM_PFCFAPI_Events_DAccountOnPositionError(self, ERRORCODE, MESSAGE):
        print("OnDAccountOnPositionError %s %s" %(ERRORCODE, MESSAGE))
        return True

    def ICOM_PFCFAPI_Events_DAccountOnUnLiquidationMainData(self, count, recordno, investorAcno, BS, ProductId,
                                                            TotalOTQTY,
                                                            RefTotalPrice, RefTotalPL, AvgMatchPrice, productKind,
                                                            Curren, RealPrice,
                                                            multiplecomno, multipleBS, multipleMatchPrice1,
                                                            multipleMatchPrice2,
                                                            PriceDiff, MultiName):
        print("筆數 %s"
              "\t目前第幾筆 %s"
              "\t帳號 %s"
              "\t買賣別 %s"
              "\t商品代碼 %s"
              "\t未平倉口數 %s"
              "\t參考現值 %s"
              "\t參考浮動損益 %s"
              "\t平均成交價 %s"
              "\t幣別 %s"
              "\t參考即時價 %s"
              "\t複式商品代碼 %s"
              "\t複式買賣別 %s"
              "\t複式第1隻腳價格 %s"
              "\t複式第2隻腳價格 %s"
              "\t價差 %s"
              "\t複式種類 %s" % (count, recordno, investorAcno, BS, ProductId, TotalOTQTY, RefTotalPrice,
                                       RefTotalPL, AvgMatchPrice, Curren, RealPrice, multiplecomno, multipleBS,
                                       multipleMatchPrice1, multipleMatchPrice2, PriceDiff, MultiName))
        return True

    def ICOM_PFCFAPI_Events_DAccountOnUnLiquidationMainError(self, ERRORCODE, MESSAGE):
        print("OnDAccountOnUnLiquidationMainError %s %s" % (ERRORCODE, MESSAGE))
        return True

    """ QUOTE events """

    def ICOM_PFCFAPI_Events_DQuoteOnTickDataBeforeTrade(self, COMMODITYID, InfoTime, MatchTime, MatchPrice, MatchBuyCnt,
                                                        MatchSellCnt,
                                                        MatchQuantity, MatchTotalQty, MatchPriceData, MatchQtyData):
        print("OnDQuoteOnTickDataBeforeTrade")
        return True

    def ICOM_PFCFAPI_Events_DQuoteOnTickDataTrade(self, COMMODITYID, InfoTime, MatchTime, MatchPrice, MatchBuyCnt,
                                                  MatchSellCnt,
                                                  MatchQuantity, MatchTotalQty, MatchPriceData, MatchQtyData):
        print("OnDQuoteOnTickDataTrade")
        return False

    def ICOM_PFCFAPI_Events_DQuoteOnDisconnected(self):
        print("OnDQuoteOnDisconnected")
        return True

    def ICOM_PFCFAPI_Events_DQuoteOnConnected(self):
        print("已連上內期行情主機事件 OnDQuoteOnConnected")
        return True

    def ICOM_PFCFAPI_Events_DQuoteOnTickDataBeforeBidOffer(self, COMMODITYID, BP1, BP2, BP3, BP4, BP5, BQ1, BQ2, BQ3,
                                                           BQ4, BQ5, SP1, SP2,
                                                           SP3, SP4, SP5, SQ1, SQ2, SQ3, SQ4, SQ5):
        print("OnDQuoteOnTickDataBeforeBidOffer")
        return True

    def ICOM_PFCFAPI_Events_DQuoteOnTickDataBidOffer(self, COMMODITYID, BP1, BP2, BP3, BP4, BP5, BQ1, BQ2, BQ3, BQ4,
                                                     BQ5, SP1, SP2,
                                                     SP3, SP4, SP5, SQ1, SQ2, SQ3, SQ4, SQ5):
        print("OnDQuoteOnTickDataBidOffer")
        return True

    def ICOM_PFCFAPI_Events_DQuoteOnIndexData(self, index_kind, index_time,
                                              index_value):
        print("OnDQuoteOnIndexData")
        return True

    def ICOM_PFCFAPI_Events_DQuoteOnTickDataHighLow(self, ProductId, dayHighPrice, dayLowPrice, showTime):
        print("OnDQuoteOnTickDataHighLow")
        return True

    """ TRADE events """

    def ICOM_PFCFAPI_Events_DTradeOnReply(self, replyData):
        print("OnDTradeOnReply: %s" % replyData)
        return True

    def ICOM_PFCFAPI_Events_DTradeOnMatch(self, matchData):
        print("OnDTradeOnMatch")
        return True

    def ICOM_PFCFAPI_Events_DTradeOnDisConnected(self):
        print("OnDTradeOnDisConnected")
        return True

    def ICOM_PFCFAPI_Events_DTradeOnConnected(self):
        print("已連上內期交易主機事件 OnDTradeOnConnected")
        return True

    def ICOM_PFCFAPI_Events_DTradeOnQueryReply(self, count, recordno, replyData):
        print("OnDTradeOnQueryReply %s %s %s" % (count, recordno, replyData))
        return True

    def ICOM_PFCFAPI_Events_DTradeOnQueryMatch(self, count, recordno, matchData):
        print("OnDTradeOnQueryMatch %s %s %s" % (count, recordno, matchData))
        return True


def setUpEventHandler():
    sink = EventSink()
    connection = comtypes.client.GetEvents(client, sink)
    PumpEvents(-1)


def login():
    print("請輸入帳號:")
    account = input("")
    print("請輸入密碼:")
    password = input("")
    print("請輸入連線ip:")
    ip_address = input("")
    client.PFCLogin(account, password, ip_address)
    selectAccount(account)


def selectAccount(login_account):
    dict = {}
    for account in client.UserOrderSet:
        dict[len(dict) + 1] = account

    global ACTNO
    if len(client.UserOrderSet) > 1:
        print(dict)
        select_account = False
        while not select_account:
            print("請依編號選擇帳號:")
            num = input()
            try:
                print("你選擇了：%s" % dict[int(num)])
                ACTNO = dict[int(num)]
                select_account = True
            except (KeyError, ValueError):
                print("無此編號，請重新輸入。")
    elif len(client.UserOrderSet) > 0:
        ACTNO = client.UserOrderSet[0]
    else:
        ACTNO = login_account


def logout():
    time.sleep(5)
    client.PFCLogout()


# 取得外期交易所資料
def PFCGetForeignEXCHANGEData():
    data = client.PFCGetForeignEXCHANGEDataPython()
    print("外期交易所資料第一筆:%s" % data[0])


# 取得外期商品資料
def PFCGetForeignSymbolData(exchange, query_type):
    data = client.PFCGetForeignSymbolDataPython(exchange, query_type)
    print("外期商品資料第一筆:%s" % data[0])


# 取得外期商品合約資料
def PFCGetForeignContractData(exchange, query_type, symbol):
    data = client.PFCGetForeignContractDataPython(exchange, query_type, symbol)
    print("外期商品合約資料第一筆:%s" % data[0])


# 外期註冊商品行情
def FQuoteRegItem():
    item = comtypes.client.CreateObject("COM_PFCFAPI.FQuoteContract")
    item.EXCHANGE = "CME"
    item.SYMBOL = "NQ"
    item.YM = "202009"
    item.CP = CPEnum.Future
    item.STRIKEPRICE = ""
    client.FQuoteRegItem(item)


# 外期反註冊商品行情
def FQuoteUnRegItem():
    item = comtypes.client.CreateObject("COM_PFCFAPI.FQuoteContract")
    item.EXCHANGE = "CME"
    item.SYMBOL = "NQ"
    item.YM = "202009"
    item.CP = CPEnum.Future
    item.STRIKEPRICE = ""
    client.FQuoteUnRegItem(item)


# 外期查詢最後報價
def FQuoteQueryItem():
    item = comtypes.client.CreateObject("COM_PFCFAPI.FQuoteContract")
    item.EXCHANGE = "CME"
    item.SYMBOL = "NQ"
    item.YM = "202009"
    item.CP = CPEnum.Future
    item.STRIKEPRICE = ""
    client.FQuoteQueryItem(item)


# 外期委託送單
def FTradeOrder():
    order = comtypes.client.CreateObject("COM_PFCFAPI.FTradeOrderObject")
    order.ACTNO = ACTNO
    order.EXCHANGE = "CME"
    order.SYMBOL = "ED"
    order.MATURITYMONTHYEAR = "202009"
    order.PUTORCALL = CPEnum.Future
    order.STRIKEPRICE = 0
    order.BS = SideEnum.Buy
    order.ORDERTYPE = OrderTypeEnum.Limit
    order.PRICE = 99.74
    order.ORDERQTY = 5
    order.TIMEINFORCE = TimeInForceEnum.ROD
    order.OPENCLOSE = OpenCloseEnum.Y
    order.DTRADE = DayTradeEnum.N
    order.NOTE = "APINoteAdd"
    client.FTradeOrder(order)

    # 複式單
    # order.SYMBOL2
    # order.MATURITYMONTHYEAR2
    # order.PUTORCALL2
    # order.STRIKEPRICE2
    # order.SIDE1
    # order.SIDE2

    # 停損
    # order.STOPPRICE


# 操作改價
def changeOrder():
    time.sleep(3)
    print("輸入改價的原委託單號:")
    order_no = input("")

    order = comtypes.client.CreateObject("COM_PFCFAPI.FTradeReplaceObject")
    order.ReplaceType = ReplaceTypeEnum.ChangePrice
    order.ACTNO = ACTNO
    order.ORDERNO = order_no
    order.ORDERTYPE = OrderTypeEnum.Limit
    order.PRICE = 99.75
    order.NOTE = "APINoteCxl"
    result = client.FTradeReplaceOrder(order)  # result FTradeORDERRESULT

    # 停損價
    # order.STOPPRICE


# 操作減量
def reduceQty():
    time.sleep(2)
    print("輸入減量的原委託單號:")
    order_no = input("")
    order = comtypes.client.CreateObject("COM_PFCFAPI.FTradeReplaceObject")
    order.ReplaceType = ReplaceTypeEnum.Decrease
    order.ACTNO = ACTNO
    order.ORDERNO = order_no
    order.ORDERQTY = 3
    order.NOTE = "APINoteDEL"
    result = client.FTradeReplaceOrder(order)  # result FTradeORDERRESULT


# 操作刪單
def deleteOrder():
    time.sleep(2)
    print("輸入刪單的原委託單號:")
    order_no = input("")
    order = comtypes.client.CreateObject("COM_PFCFAPI.FTradeReplaceObject")
    order.ReplaceType = ReplaceTypeEnum.Cancel
    order.ACTNO = ACTNO
    order.ORDERNO = order_no
    order.NOTE = "APINoteCxl"
    result = client.FTradeReplaceOrder(order)  # result FTradeORDERRESULT


# 查詢委託回報
def FTradeQueryReply():
    num_of_query = "10"
    network_id_start = ""
    network_id_end = ""
    begin_order_time = ""
    end_order_time = ""
    client.FTradeQueryReply(ACTNO, num_of_query, network_id_start, network_id_end, begin_order_time, end_order_time)


# 查詢成交回報
def FTradeQueryMatch():
    num_of_query = "10"
    network_id_start = ""
    network_id_end = ""
    begin_order_time = ""
    end_order_time = ""
    client.FTradeQueryMatch(ACTNO, num_of_query, network_id_start, network_id_end, begin_order_time, end_order_time)


# 查詢保證金
def FAccountGetMargin():
    print("FAccountGetMargin")
    currency = "NTD"
    client.FAccountGetMargin(ACTNO, currency)


# 查詢即時部位
def FAccountGetPosition():
    productId = ""
    client.FAccountGetPosition(ACTNO, productId)


# 查詢未平倉
def FAccountGetUnLiquidationDetail():
    client.FAccountGetUnLiquidationDetail(ACTNO)


def exitProgram():
    try:
        os._exit(0)
    except:
        print('exit')


if __name__ == '__main__':
    print("Python version")
    print(sys.version)

    # setup event handler
    thread = Thread(target=setUpEventHandler)
    thread.start()

    login()

    # 取得外期交易所資料
    PFCGetForeignEXCHANGEData()  # return PFCFAPI.PFCFAPI.EXCHANGEDATA

    # 取得外期商品資料
    PFCGetForeignSymbolData("", "")  # return PFCFAPI.PFCFAPI.SYMBOLDATA

    # 取得外期商品合約資料
    PFCGetForeignContractData("", "", "")  # return PFCFAPI.PFCFAPI.CONTRACTDATA

    # 外期註冊商品行情
    FQuoteRegItem()

    # 沒作用
    # 外期查詢最後報價
    # FQuoteQueryItem()

    # 外期委託送單
    FTradeOrder()

    # 操作改價
    # changeOrder()

    # 操作減量
    reduceQty()

    # 操作刪單
    deleteOrder()

    # 查詢委託回報
    FTradeQueryReply()

    # 查詢成交回報
    FTradeQueryMatch()

    # 查詢保證金
    FAccountGetMargin()

    # 查詢即時部位
    FAccountGetPosition()

    # 查詢未平倉
    FAccountGetUnLiquidationDetail()

    logout()

    exitProgram()
