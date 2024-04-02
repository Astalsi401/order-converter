import os
import re
import pandas as pd
import logging
from datetime import datetime as dt
from pickle import load

os.chdir(os.path.dirname(os.path.abspath(__file__)))
pd.set_option('display.max.columns', None)
result = 'import'


def readXlsx(path, converters: dict):
    '''讀取excel，若檔案不存在則回傳空dataframe'''
    return pd.read_excel(path, converters=converters) if os.path.isfile(path) else pd.DataFrame()


def getFilesName(path, ext=None):
    '''抓取資料夾內檔案名稱, ext指定副檔名'''
    if ext == None:
        return [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
    else:
        return [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f)) and f'.{ext}' in f]


class SourceFiles:
    def __init__(self) -> None:
        self.yahoo_mall = 'yahoo購物中心宅配(管制明文).xlsx'
        self.yahoo_shop_h = 'yahoo商城宅配.xlsx'
        self.yahoo_shop_s = 'yahoo商城店配.xlsx'
        self.shopee = 'shopee店配宅配(管制明文).xlsx'
        self.shopline = 'shopline店配宅配(管制明文).xlsx'
        self.rakuten = 'rakuten店配宅配(管制明文).xlsx'


class ColumnType:
    def __init__(self) -> None:
        self.yahoo = 'yahoo'
        self.shopee = 'shopee'
        self.shopline = 'shopline'
        self.rakuten = 'rakuten'


class OutputColumns:
    def __init__(self, fr) -> None:
        self.fr = fr
        self.code = '外站訂單編號'
        self.site = '外站名稱'
        self.customer = '收件人'
        self.postCode = '郵遞區號'
        self.address = '地址'
        self.tel = '電話'
        self.cel = '行動電話'
        self.date = '訂單成立日期'
        self.productCode = '產品編號'
        self.product = '產品名稱'
        self.number = '購買數量'
        self.price = '商品總金額'
        self.payCode = '付款代號'
        self.pay = '付款方式'
        self.manufacture = '廠商名稱'
        self.purchasePrice = '進貨價'
        self.purchaseSubtotal = '進貨小計'
        self.warehouse = '預設倉庫'
        self.tradeCode = '交易序號'
        self.subtotal = '訂單金額'
        self.tally = '入庫出貨撿貨費'
        self.orderFee = '訂單處理費'
        self.ship = '運費\n(箱子+包材+運費)'
        self.profit = '利潤'
        self.profitPc = '利潤百分比'
        self.pm = '負責PM'
        self.note = '備註'
        self.serviceFee = '隱碼服務費'
        self.send = '寄送方式'
        self.priceImport = '待定'
        if self.fr == ColumnType().yahoo:
            self.cashFee = '金流費用'
            self.fee = '成交手續費\n(含購物車費用)'
            self.profitDenominator = self.subtotal
            self.profitCols = [self.purchaseSubtotal, self.fee, self.cashFee, self.tally, self.orderFee, self.ship]
            self.finCols = [self.code, self.site, self.customer, self.postCode, self.address, self.tel, self.cel, self.date, self.productCode, self.product, self.number, self.price, self.payCode, self.pay, self.manufacture, self.purchasePrice, self.purchaseSubtotal, self.warehouse, self.tradeCode, self.fee, self.cashFee, self.tally, self.orderFee, self.ship, self.profit, self.profitPc, self.pm, self.note]
            self.rename = {
                '訂單編號': self.code,
                '收件人姓名': self.customer,
                '收件人郵遞區號': self.postCode,
                '收件人地址': self.address,
                '收件人電話(日)': self.tel,
                '收件人行動電話': self.cel,
                '轉單日期': self.date,
                '店家商品料號': self.productCode,
                '數量': self.number,
                '付款別': self.pay,
                '交易序號': self.tradeCode
            }
        elif self.fr == ColumnType().shopee:
            self.discount = '蝦幣回饋券'
            self.fee = '交易手續費12%\n(成交、活動、金流)'
            self.profitDenominator = self.subtotal
            self.profitCols = [self.purchaseSubtotal, self.fee, self.discount, self.tally, self.orderFee, self.ship]
            self.finCols = [self.code, self.site, self.customer, self.postCode, self.address, self.tel, self.cel, self.date, self.productCode, self.product, self.number, self.price, self.payCode, self.pay, self.manufacture, self.purchasePrice, self.purchaseSubtotal, self.warehouse, self.tradeCode, self.subtotal, self.fee, self.discount, self.serviceFee, self.tally, self.orderFee, self.ship, self.profit, self.profitPc, self.pm, self.note]
            self.rename = {
                '訂單編號': self.code,
                '收件者姓名': self.customer,
                '郵遞區號': self.postCode,
                '收件地址': self.address,
                '蝦皮專線和包裹查詢碼 \n(請複製下方完整編號提供給您配合的物流商當做聯絡電話)': self.cel,
                '訂單成立日期': self.date,
                '商品選項貨號': self.productCode,
                '數量': self.number,
                '付款方式': self.pay,
                '賣家蝦幣回饋券': self.discount
            }
        elif self.fr == ColumnType().shopline:
            self.fee = '交易手續費2.8%\n(成交、金流)'
            self.profitDenominator = self.subtotal
            self.discount = '優惠折扣'
            self.customDiscount = '自訂折扣合計'
            self.coupon = '折抵購物金'
            self.reward = '點數折現'
            self.profitCols = [self.purchaseSubtotal, self.fee, self.tally, self.orderFee, self.ship]
            self.shipCols = [self.site, self.code, self.customer, self.postCode, self.address, self.tel, self.cel, self.productCode, self.date, self.price]
            self.finCols = [self.code, self.site, self.customer, self.postCode, self.address, self.tel, self.cel, self.date, self.productCode, self.product, self.number, self.price, self.discount, self.customDiscount, self.coupon, self.reward, self.payCode, self.pay, self.manufacture, self.purchasePrice, self.purchaseSubtotal, self.warehouse, self.tradeCode, self.subtotal, self.fee, self.tally, self.orderFee, self.ship, self.profit, self.profitPc, self.pm, self.note]
            self.rename = {
                '訂單號碼': self.code,
                '收件人': self.customer,
                '郵政編號（如適用)': self.postCode,
                '完整地址': self.address,
                '電話號碼': self.tel,
                '收件人電話號碼': self.cel,
                '訂單日期': self.date,
                '商品貨號': self.productCode,
                '數量': self.number,
                '付款方式': self.pay,
                '送貨方式': self.send
            }
        elif self.fr == ColumnType().rakuten:
            self.priceImport = '匯入輔翼金額'
            self.discount = '店家優惠券'
            self.fee = '手續費'  # 目前未使用
            self.profitDenominator = self.subtotal
            self.profitCols = [self.ship]
            self.shipCols = [self.date, self.code, self.customer, self.cel, self.postCode, self.address, self.productCode, self.product, self.number, self.price, self.payCode, self.pay, self.discount]
            self.finCols = [self.date, self.code, self.customer, self.cel, self.postCode, self.address, self.productCode, self.product, self.number, self.price, self.ship, self.priceImport, self.payCode, self.pay, self.discount]
            self.rename = {
                '訂單日期': self.date,
                '訂單號碼': self.code,
                '收件人姓名': self.customer,
                '收件人的電話號碼': self.cel,
                '目的地郵遞區號': self.postCode,
                '配送地址': self.address,
                '商品管理編號 (SKU)': self.productCode,
                '商品名稱': self.product,
                '購買數量': self.number,
                '商品價格': self.price,
                '配送方式': self.send,
                '訂單與運費總和': self.subtotal
            }
        # 需四捨五入的欄位
        self.roundCols = [self.price, self.purchasePrice, self.purchaseSubtotal, self.fee, self.subtotal, self.profit, self.priceImport]


class Price:
    def __init__(self, sum: list, col: str = None) -> None:
        self.sum = sum
        self.col = col


class FeeRate:
    def __init__(self, pc: float, add=0) -> None:
        self.pc = pc
        self.add = add


class Converter:
    def __init__(self, fr: list, cov: dict, oc: dict, price: Price, timeFmt: str, fileName: str, feeRate: FeeRate = None) -> None:
        self.tmp = 'tmp'
        self.count = 'count'
        self.fr = fr
        self.oc = oc
        self.cov = cov
        self.price = price
        self.timeFmt = timeFmt
        self.fileName = f'{result}/{fileName}.xlsx'
        settings = load(open(f'設定/settings.pkl', 'rb'))
        self.feeRate = feeRate if feeRate else FeeRate(settings[fileName]['feeRate']['rate'], settings[fileName]['feeRate']['add'])
        self.payCode = {
            1: [{self.oc.pay: ['銀行轉帳', '蝦皮錢包', '線上支付', 'ATM/銀行轉帳', 'ATM', '全家繳費']}],
            3: [{self.oc.pay: ['貨到付款', '現付', '7-11門市取貨付款']}],
            4: [{self.oc.pay: ['信用卡', '信用卡分期付款', 'LINE Pay', '信用卡付款', '信用卡一次', '分期付款', '街口支付']}],
            6: [{self.oc.pay: ['7-11']}, {self.oc.send: ['7-ELEVEN'], self.oc.pay: ['貨到付款']}, {self.oc.send: ['7-11 取貨 (到店付款)', '全家取貨 (到店付款)'], self.oc.pay: ['貨到付款', '貨到付款']}],
        }
        self.ship = {
            75: [{self.oc.site: ['yahoo商城店配']}, {self.oc.send: ['7-ELEVEN', '7-11 取貨 (到店付款)', '全家取貨 (到店付款)']}],
            140: [{self.oc.site: ['yahoo商城宅配', 'yahoo購物中心宅配']}, {self.oc.send: ['賣家宅配', '宅配']}],
        }
        self.serviceFee = {
            0: [{self.oc.site: ['yahoo商城店配']}, {self.oc.send: ['7-ELEVEN', '7-11 取貨 (到店付款)', '全家取貨 (到店付款)']}],
            10: [{self.oc.site: ['yahoo商城宅配', 'yahoo購物中心宅配']}, {self.oc.send: ['賣家宅配', '宅配']}],
        }
        # 如有複數來源檔案須將檔案合併
        self.df = self.concatFr()
        if self.df.empty:
            return None
        logging.info(f'正在轉檔：{self.fr}')
        # 付款代號
        self.df = self.multiCondition(self.payCode, [self.oc.payCode])
        # 刪除shopline已取消\非貨到付款且未付款的訂單
        if self.oc.fr == ColumnType().shopline:
            self.df = self.df.query(f'訂單狀態 != "已取消"')
            self.df = self.df[~((self.df[self.oc.payCode] != 6) & (self.df[self.oc.payCode] != 3) & (self.df['付款狀態'] == '未付款'))]
        # 辨別是否為該訂單第一件商品
        self.df[self.count] = self.df.groupby(self.oc.code).cumcount()
        # 訂單編號
        self.df[self.oc.code] = self.df[self.oc.code].str.replace(r'#', '', regex=True)
        # 郵遞區號取前三碼
        self.df[self.oc.postCode] = self.df[self.oc.postCode].fillna('').apply(lambda x: x[:3])
        # 更改訂單成立日期格式
        self.df[self.oc.date] = self.df[self.oc.date].fillna('').astype(str).apply(lambda x: dt.strptime(x, self.timeFmt).strftime('%Y%m%d'))
        # 付款方式
        self.df[self.oc.pay] = self.df[self.oc.pay].str.replace(r'\([^\(|\)]*\)', '', regex=True)
        # 商品總金額
        self.df[self.oc.price] = self.df[self.price.sum].sum(axis=1)
        # 替換空白電話號碼為'****'
        self.df.loc[self.df[self.oc.cel].isna(), self.oc.cel] = '****'
        # 如果在後續金額計算中需要把商品價格*購買數量
        if self.price.col:
            self.df[self.tmp] = self.df[self.price.col] * self.df[self.oc.number]
        # shopline商品總金額要減掉運費
        if self.oc.fr in [ColumnType().shopline]:
            self.df[self.oc.price] = self.df[self.oc.price] - self.df['運費']
        if self.oc.fr in [ColumnType().shopee]:
            self.df[self.oc.cel] = self.df[self.oc.cel].replace(r'#\d$', '', regex=True)
        if self.oc.fr in [ColumnType().shopee, ColumnType().shopline]:
            self.df[self.oc.price] = self.df.groupby(self.oc.code)[self.oc.price].transform('first') / self.df.groupby(self.oc.code)[self.tmp].transform(lambda x: x.sum()) * self.df[self.tmp]
        # rakuten 匯入輔翼金額，不確定未來是否要整合至訂單金額
        if self.oc.fr in [ColumnType().rakuten]:
            self.df[self.tmp] = self.df.groupby(self.oc.code)[[self.tmp]].transform(lambda x: x.sum()).sum(axis=1)
            self.df[self.oc.priceImport] = (1 - (self.df[self.oc.discount] / self.df[self.tmp])) * self.df[self.oc.price] * self.df[self.oc.number]
        # shopline, rakuten運費
        if self.oc.fr in [ColumnType().shopline, ColumnType().rakuten]:
            ship = self.df.loc[self.df['運費'] > 0].copy()
            ship[self.oc.productCode] = '888888888'
            ship[self.oc.price] = ship['運費']
            ship = ship[self.oc.shipCols]
            self.df = pd.concat([self.df, ship]).sort_values(by=[self.oc.date, self.oc.code])
        # 合併商品總表中的資訊並計算撿貨費
        self.df = self.productDetail(self.df)
        # 撿貨費
        self.df[self.oc.tally] = self.df['撿貨數量'] * self.df[self.oc.number] * 7.5
        # 進貨小計
        self.df[self.oc.purchaseSubtotal] = self.df[self.oc.purchasePrice] * self.df[self.oc.number]
        # 成交手續費
        self.df[self.oc.fee] = self.df[self.oc.price] * self.feeRate.pc + self.feeRate.add
        if self.oc.fr in [ColumnType().yahoo]:
            # 金流費用
            self.df.loc[self.df[self.oc.payCode] == 1, self.oc.cashFee] = 0
            self.df.loc[self.df[self.oc.payCode] == 4, self.oc.cashFee] = self.df[self.oc.price] * 0.02
            self.df.loc[self.df[self.oc.pay] == '全家繳費', self.oc.cashFee] = 15
            self.df.loc[self.df[self.oc.pay] == '7-11', self.oc.cashFee] = 48
            # 訂單金額
            self.df[self.oc.subtotal] = self.df.groupby(self.oc.code)[self.price.sum].transform(lambda x: x.sum()).sum(axis=1)
        elif self.oc.fr in [ColumnType().shopee, ColumnType().shopline]:
            # 訂單金額
            self.df[self.oc.subtotal] = self.df[self.price.sum].sum(axis=1)
        # 訂單處理費
        self.df[self.oc.orderFee] = (self.df.groupby(self.oc.code)[self.oc.code].transform("count") - 1) * 10 + 26
        # 運費
        self.df = self.multiCondition(self.ship, [self.oc.ship])
        # shopee隱碼服務費
        self.df = self.multiCondition(self.serviceFee, [self.oc.serviceFee])
        # 依倉庫調整撿貨費、訂單處理費、運費
        self.df.loc[self.df[self.oc.warehouse].fillna('').str.contains(r'^(?:原廠出貨|公司倉)$', regex=True), [self.oc.tally, self.oc.orderFee]] = 0
        self.df.loc[self.df[self.oc.warehouse] == '原廠出貨', self.oc.ship] = 0
        # 如果不是第一件商品，則'運費','訂單處理費','隱碼服務費'為0
        self.df.loc[self.df[self.count] != 0, [self.oc.ship, self.oc.orderFee, self.oc.serviceFee]] = 0
        # 利潤
        self.df[self.oc.subtotal] = self.df.groupby(self.oc.code)[self.oc.subtotal].transform('first')
        self.df[self.oc.profit] = self.df[self.oc.subtotal] - self.df.groupby(self.oc.code)[self.oc.profitCols].transform(lambda x: x.sum()).sum(axis=1)
        # 利潤百分比
        self.df[self.oc.profitPc] = (self.df[self.oc.profit] / self.df[self.oc.profitDenominator] * 100).round(2).astype(str) + '%'
        # 補齊需要的欄位
        self.df = self.addColumns()
        # 如果不是第一件商品，則'訂單金額','利潤','利潤百分比'為0
        self.df.loc[self.df[self.count] != 0, [self.oc.subtotal, self.oc.profit, self.oc.profitPc]] = 0
        # 價格四捨五入至整數
        for col in self.oc.roundCols:
            if col in self.df.columns:
                self.df[col] = self.df[col].fillna(0).round(0)
        # 匯出需要的欄位
        self.df = self.addColumns()[self.oc.finCols]

    def concatFr(self):
        dfList = []
        for file in self.fr:
            df = readXlsx(f'待轉檔/{file}', converters=self.cov)
            df[self.oc.site] = re.sub(r'(.xlsx)$', '', file)
            if file == SourceFiles().yahoo_shop_s and not df.empty:
                df = df.rename(columns={'收件人電話': '收件人電話(日)', '轉單日': '轉單日期'})
                df['收件人地址'] = df['超商類型'] + df['收件人地址']
                df['付款別'] = df['超商類型']
                df['收件人郵遞區號'] = ''
            dfList.append(df)
        return pd.concat(dfList).rename(columns=self.oc.rename)

    def multiCondition(self, conditions: dict, cols: list):
        # 將篩選條件作為新的df進行left_merge，並將欄位col取代為key值
        for key, dfs in conditions.items():
            for df in dfs:
                if set(df).issubset(set(self.df.columns)):
                    self.df = self.df.merge(pd.DataFrame(df), how='left', indicator=True)
                    self.df.loc[self.df['_merge'] == 'both', cols] = key
                    self.df.drop(columns=['_merge'], inplace=True)
        return self.df

    def productDetail(self, df):
        '''合併商品總表資料'''
        df.drop(columns=[col for col in [self.oc.product] if col in df.columns], inplace=True)
        d = pd.read_excel('設定/商品總表(管制明文).xlsx', converters={'商品代號': str})
        d.loc[~d['進貨價_活動'].isna(), '進貨價'] = d['進貨價_活動']
        return df.merge(d.rename(columns={'商品代號': self.oc.productCode, '商品名稱': self.oc.product, '廠商名稱': self.oc.manufacture, '進貨價': self.oc.purchasePrice, '預設倉庫': self.oc.warehouse, '負責PM': self.oc.pm}), on=self.oc.productCode, how='left')

    def addColumns(self):
        '''依據expCol補齊所需的欄位'''
        for col in self.oc.finCols:
            if col not in self.df.columns:
                self.df[col] = ''
        return self.df

    def to_excel(self):
        logging.info(f'{self.fileName} saved!')
        self.df.to_excel(self.fileName, index=False)


def main():
    for f in getFilesName(result, 'xlsx'):
        try:
            os.remove(f'{result}/{f}')
        except:
            pass
    yahoo_shop = Converter(
        fr=[SourceFiles().yahoo_shop_h, SourceFiles().yahoo_shop_s],
        cov={'交易序號': str, '訂單編號': str, '商品編號': str, '收件人電話(日)': str, '收件人行動電話': str, '收件人電話': str, '收件人郵遞區號': str, '轉單日期': str},
        oc=OutputColumns(ColumnType().yahoo),
        price=Price(['金額小計', '超贈點折抵金額', '行銷補助金額']),
        timeFmt='%Y/%m/%d %H:%M',
        fileName='yahoo商城'
    )
    yahoo_mall = Converter(
        fr=[SourceFiles().yahoo_mall],
        cov={'交易序號': str, '訂單編號': str, '店家商品料號': str, '收件人電話(日)': str, '收件人行動電話': str, '收件人郵遞區號': str, '轉單日期': str},
        oc=OutputColumns(ColumnType().yahoo),
        price=Price(['金額小計', '超贈點折抵金額', '行銷補助金額']),
        timeFmt='%Y/%m/%d %H:%M',
        fileName='yahoo購物中心'
    )
    shopee = Converter(
        fr=[SourceFiles().shopee],
        cov={'訂單編號': str, '商品選項貨號': str, '收件者電話': str, '取件門市店號': str, '郵遞區號': str, '訂單成立日期': str},
        oc=OutputColumns(ColumnType().shopee),
        price=Price(['買家總支付金額', '蝦幣折抵', '銀行信用卡活動折抵', '優惠券'], '商品活動價格'),
        timeFmt='%Y-%m-%d %H:%M',
        fileName='shopee'
    )
    shopline = Converter(
        fr=[SourceFiles().shopline],
        cov={'訂單號碼': str, '郵政編號（如適用)': str, '電話號碼': str, '收件人電話號碼': str, '訂單成立日期': str, '商品貨號': str, '全家服務編號 / 7-11 店號': str},
        oc=OutputColumns(ColumnType().shopline),
        price=Price(['付款總金額'], '商品結帳價'),
        timeFmt='%Y-%m-%d %H:%M:%S',
        fileName='shopline'
    )
    rakuten = Converter(
        fr=[SourceFiles().rakuten],
        cov={'訂單日期': str, '訂單號碼': str, '收件人的電話號碼': str, '目的地郵遞區號': str, '商品管理編號 (SKU)': str},
        oc=OutputColumns(ColumnType().rakuten),
        price=Price(['商品總金額'], '商品總金額'),
        timeFmt='%Y-%m-%d %H:%M:%S',
        fileName='rakuten'
    )
    for cov in [shopee, shopline, yahoo_mall, yahoo_shop, rakuten]:
        if not cov.df.empty:
            cov.to_excel()


if __name__ == '__main__':
    logFile = '設定/run.log'
    logging.basicConfig(format='%(asctime)s %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S', level=logging.INFO, handlers=[logging.FileHandler(logFile), logging.StreamHandler()])
    try:
        main()
    except:
        logging.exception(f'錯誤訊息已處存至 {logFile}')
    input('按Enter繼續...')
