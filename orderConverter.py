import os
import re
import pandas as pd
import logging
from io import BytesIO
from msoffcrypto import OfficeFile, exceptions
from datetime import datetime as dt
from pickle import load

os.chdir(os.path.dirname(os.path.abspath(__file__)))
pd.set_option('display.max.columns', None)
result = 'import'


def get_password(setting_key: str) -> str:
    setting_pkl = '設定/settings.pkl'
    if not os.path.exists(setting_pkl):
        raise FileNotFoundError('請先執行setting.py，並儲存預設設定檔')
    try:
        password = load(open(setting_pkl, 'rb'))[setting_key]['password']
        return None if password == '' else password
    except KeyError:
        raise KeyError('請先執行setting.py，並設定檔案密碼')


def read_xlsx(path: str, converters: dict[str, classmethod], password=None) -> pd.DataFrame:
    '''讀取excel，若檔案不存在則回傳空dataframe'''
    if not os.path.isfile(path):
        return pd.DataFrame()
    if password:
        try:
            data = BytesIO()
            office_file = OfficeFile(open(path, 'rb'))
            office_file.load_key(password=password)
            office_file.decrypt(data)
        except exceptions.DecryptionError as e:
            if str(e) == 'Unencrypted document':
                logging.warning(f'{path} 不需要密碼')
                password = None
            else:
                logging.error(f'{path} 解密失敗')
    return pd.read_excel(data if password else path, converters=converters)


def get_files_name(path: str, ext=None) -> list[str]:
    '''抓取資料夾內檔案名稱, ext指定副檔名'''
    return [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))] if ext == None else [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f)) and f'.{ext}' in f]


class SourceFiles:
    class Source:
        def __init__(self, file: str, setting: str) -> None:
            self.file = file
            self.setting = setting
            self.site = re.sub(r'\(管制明文\)|\.xlsx$', '', self.file)

    def __init__(self) -> None:
        self.yahoo_mall = self.Source('yahoo購物中心宅配(管制明文).xlsx', 'yahoo購物中心')
        self.yahoo_shop_h = self.Source('yahoo商城宅配.xlsx', 'yahoo商城')
        self.yahoo_shop_s = self.Source('yahoo商城店配.xlsx', 'yahoo商城')
        self.shopee = self.Source('shopee店配宅配(管制明文).xlsx', 'shopee')
        self.rakuten = self.Source('rakuten店配宅配(管制明文).xlsx', 'rakuten')
        self.shopline = self.Source('shopline店配宅配(管制明文).xlsx', 'shopline')


class ColumnType:
    def __init__(self) -> None:
        self.yahoo = 'yahoo'
        self.shopee = 'shopee'
        self.shopline = 'shopline'
        self.rakuten = 'rakuten'


class OutputColumns:
    def __init__(self, fr: str) -> None:
        self.fr = fr
        self.code = '外站訂單編號'
        self.site = '外站名稱'
        self.customer = '收件人'
        self.post_code = '郵遞區號'
        self.address = '地址'
        self.tel = '電話'
        self.cel = '行動電話'
        self.date = '訂單成立日期'
        self.product_code = '產品編號'
        self.product = '產品名稱'
        self.number = '購買數量'
        self.price = '商品總金額'
        self.pay_code = '付款代號'
        self.pay = '付款方式'
        self.manufacture = '廠商名稱'
        self.purchase_price = '進貨價'
        self.purchase_subtotal = '進貨小計'
        self.warehouse = '預設倉庫'
        self.trade_code = '交易序號'
        self.subtotal = '訂單金額'
        self.tally = '入庫出貨撿貨費'
        self.order_fee = '訂單處理費'
        self.ship = '運費\n(箱子+包材+運費)'
        self.profit = '利潤'
        self.profit_pc = '利潤百分比'
        self.pm = '負責PM'
        self.note = '備註'
        self.service_fee = '隱碼服務費'
        self.send = '寄送方式'
        self.cash_fee = '金流費用'
        self.delivery_fee = '物流交寄使用費$48'
        self.point = '點數成本負擔'
        if self.fr == ColumnType().yahoo:
            self.discount = '超贈點點數'
            self.fee = '成交手續費\n(含購物車費用)'
            self.profit_denominator = self.subtotal
            self.profit_cols = [self.purchase_subtotal, self.fee, self.cash_fee, self.tally, self.order_fee, self.ship]
            self.fin_cols = [self.code, self.site, self.customer, self.post_code, self.address, self.tel, self.cel, self.date, self.product_code, self.product, self.number, self.price, self.pay_code, self.pay, self.manufacture, self.purchase_price, self.purchase_subtotal, self.warehouse, self.trade_code, self.fee, self.cash_fee, self.tally, self.order_fee, self.ship, self.profit, self.profit_pc, self.pm, self.note]
            self.rename = {
                '訂單編號': self.code,
                '收件人姓名': self.customer,
                '收件人郵遞區號': self.post_code,
                '收件人地址': self.address,
                '收件人電話(日)': self.tel,
                '收件人行動電話': self.cel,
                '轉單日期': self.date,
                '店家商品料號': self.product_code,
                '數量': self.number,
                '付款別': self.pay,
                '交易序號': self.trade_code
            }
        elif self.fr == ColumnType().shopee:
            self.discount = '蝦幣回饋券'
            self.fee = '交易手續費12%\n(成交、活動、金流)'
            self.profit_denominator = self.subtotal
            self.profit_cols = [self.purchase_subtotal, self.fee, self.discount, self.tally, self.order_fee, self.ship]
            self.fin_cols = [self.code, self.site, self.customer, self.post_code, self.address, self.tel, self.cel, self.date, self.product_code, self.product, self.number, self.price, self.pay_code, self.pay, self.manufacture, self.purchase_price, self.purchase_subtotal, self.warehouse, self.trade_code, self.subtotal, self.fee, self.discount, self.service_fee, self.tally, self.order_fee, self.ship, self.profit, self.profit_pc, self.pm, self.note]
            self.rename = {
                '訂單編號': self.code,
                '收件者姓名': self.customer,
                '郵遞區號': self.post_code,
                '收件地址': self.address,
                '蝦皮專線和包裹查詢碼 \n(請複製下方完整編號提供給您配合的物流商當做聯絡電話)': self.cel,
                '訂單成立日期': self.date,
                '商品選項貨號': self.product_code,
                '數量': self.number,
                '付款方式': self.pay,
                '賣家蝦幣回饋券': self.discount
            }
        elif self.fr == ColumnType().shopline:
            self.fee = '交易手續費2.8%\n(成交、金流)'
            self.profit_denominator = self.subtotal
            self.discount = '優惠折扣'
            self.custom_discount = '自訂折扣合計'
            self.coupon = '折抵購物金'
            self.reward = '點數折現'
            self.profit_cols = [self.purchase_subtotal, self.fee, self.tally, self.order_fee, self.ship]
            self.ship_cols = [self.code, self.site, self.customer, self.post_code, self.address, self.tel, self.cel, self.date, self.product_code, self.price]
            self.fin_cols = [self.code, self.site, self.customer, self.post_code, self.address, self.tel, self.cel, self.date, self.product_code, self.product, self.number, self.price, self.discount, self.custom_discount, self.coupon, self.reward, self.pay_code, self.pay, self.manufacture, self.purchase_price, self.purchase_subtotal, self.warehouse, self.trade_code, self.subtotal, self.fee, self.tally, self.order_fee, self.ship, self.profit, self.profit_pc, self.pm, self.note]
            self.rename = {
                '訂單號碼': self.code,
                '收件人': self.customer,
                '郵政編號（如適用)': self.post_code,
                '完整地址': self.address,
                '電話號碼': self.tel,
                '收件人電話號碼': self.cel,
                '訂單日期': self.date,
                '商品貨號': self.product_code,
                '數量': self.number,
                '付款方式': self.pay,
                '送貨方式': self.send
            }
        elif self.fr == ColumnType().rakuten:
            self.discount = '店家優惠券'
            self.fee = '成交\n手續費5.3%+$3'
            self.cash_fee = '金流費用\n(2%-2.2%)'
            self.tally = '入庫出貨撿貨費\n(含併件費)'
            self.order_fee = '訂單\n處理費$26'
            self.ship = '運費\n(箱子+包材+運費)\n(預設-宅配140、店配75)'
            self.profit_denominator = self.subtotal
            self.profit_cols = [self.purchase_subtotal, self.fee, self.cash_fee, self.delivery_fee, self.point, self.tally, self.order_fee, self.ship]
            self.ship_cols = [self.code, self.site, self.customer, self.post_code, self.address, self.cel, self.date, self.product_code, self.price, self.pay_code, self.pay]
            self.fin_cols = [self.code, self.site, self.customer, self.post_code, self.address, self.tel, self.cel, self.date, self.product_code, self.product, self.number, self.price, self.pay_code, self.pay, self.manufacture, self.purchase_price, self.purchase_subtotal, self.warehouse, self.fee, self.cash_fee, self.delivery_fee, self.point, self.tally, self.order_fee, self.ship, self.profit, self.profit_pc, self.pm]
            self.rename = {
                '訂單號碼': self.code,
                '收件人姓名': self.customer,
                '目的地郵遞區號': self.post_code,
                '配送地址': self.address,
                '收件人的電話號碼': self.cel,
                '訂單日期': self.date,
                '商品管理編號 (SKU)': self.product_code,
                '商品名稱': self.product,
                '購買數量': self.number,
                '訂單與運費總和': self.price,
                '配送方式': self.send,
                '商家獎勵顧客的點數總和': self.point
            }
        # 需四捨五入的欄位
        self.round_cols = [self.price, self.purchase_price, self.purchase_subtotal, self.fee, self.cash_fee, self.delivery_fee, self.point, self.subtotal, self.profit]


class Price:
    def __init__(self, sum: list[str], col: str = None) -> None:
        self.sum = sum
        self.col = col


class FeeRate:
    def __init__(self, pc: float, add=0) -> None:
        self.pc = pc
        self.add = add


class Converter:
    def __init__(self, fr: list[SourceFiles.Source], cov: dict[str, classmethod], oc: OutputColumns, price: Price, time_fmt: str, file_name: str, feeRate: FeeRate = None) -> None:
        self.tmp = 'tmp'
        self.count = 'count'
        self.fr = fr
        self.oc = oc
        self.cov = cov
        self.price = price
        self.time_fmt = time_fmt
        self.file_name = f'{result}/{file_name}.xlsx'
        settings = load(open(f'設定/settings.pkl', 'rb'))
        self.feeRate = feeRate if feeRate else FeeRate(settings[file_name]['feeRate']['rate'], settings[file_name]['feeRate']['add'])
        self.pay_code = {
            1: [{self.oc.pay: ['銀行轉帳', '蝦皮錢包', '線上支付', 'ATM/銀行轉帳', 'ATM', '全家繳費']}],
            3: [{self.oc.pay: ['貨到付款', '現付', '7-11門市取貨付款']}],
            4: [{self.oc.pay: ['信用卡', '信用卡分期付款', 'LINE Pay', '信用卡付款', '信用卡一次', '分期付款', '街口支付']}],
            6: [{self.oc.pay: ['7-11', '7-11門市取貨付款']}, {self.oc.send: ['7-ELEVEN'], self.oc.pay: ['貨到付款']}, {self.oc.send: ['7-11 取貨 (到店付款)', '全家取貨 (到店付款)'], self.oc.pay: ['貨到付款', '貨到付款']}],
        }
        self.ship = {
            75: [{self.oc.send: ['7-ELEVEN', '7-11 取貨 (到店付款)', '全家取貨 (到店付款)', '全家門市取貨', '7-11門市取貨']}],
            140: [{self.oc.site: [SourceFiles().yahoo_mall.site]}, {self.oc.send: ['賣家宅配', '宅配', '常溫宅配(倉儲中心)']}],
        }
        self.service_fee = {
            0: [{self.oc.send: ['7-ELEVEN', '7-11 取貨 (到店付款)', '全家取貨 (到店付款)']}],
            10: [{self.oc.send: ['賣家宅配', '宅配']}],
        }
        self.delivery_fee = {
            48: [{self.oc.send: ['全家門市取貨', '7-11門市取貨']}]
        }
        self.cash_fee = {
            0: [{self.oc.pay: ['銀行轉帳', '蝦皮錢包', '線上支付', 'ATM/銀行轉帳', 'ATM', '全家繳費']}],
            0.02: [{self.oc.pay: ['信用卡', '信用卡分期付款', '信用卡付款', '信用卡一次', '分期付款', '街口支付']}],
            0.022: [{self.oc.pay: ['LINE Pay']}],
            15: [{self.oc.pay: ['全家繳費']}],
            48: [{self.oc.pay: ['7-11']}],
        }

    def concat_fr(self) -> pd.DataFrame:
        return pd.concat([read_xlsx(f'待轉檔/{file.file}', converters=self.cov, password=get_password(file.setting)).assign(**{self.oc.site: file.site}) for file in self.fr]).rename(columns=self.oc.rename)

    def multi_condition(self, conditions: dict[str, list[dict[str, list]]], cols: list[str]) -> None:
        # 將篩選條件作為新的df進行left_merge，並將欄位col取代為key值
        for key, dfs in conditions.items():
            for df in dfs:
                if set(df).issubset(set(self.df.columns)):
                    self.df = self.df.merge(pd.DataFrame(df), how='left', indicator=True)
                    self.df.loc[self.df['_merge'] == 'both', cols] = key
                    self.df.drop(columns=['_merge'], inplace=True)

    def product_detail(self) -> pd.DataFrame:
        '''合併商品總表資料'''
        self.df.drop(columns=[col for col in [self.oc.product] if col in self.df.columns], inplace=True)
        d = pd.read_excel('設定/商品總表(管制明文).xlsx', converters={'商品代號': str})
        d.loc[~d['進貨價_活動'].isna(), '進貨價'] = d['進貨價_活動']
        return self.df.merge(d.rename(columns={'商品代號': self.oc.product_code, '商品名稱': self.oc.product, '廠商名稱': self.oc.manufacture, '進貨價': self.oc.purchase_price, '預設倉庫': self.oc.warehouse, '負責PM': self.oc.pm}), on=self.oc.product_code, how='left')

    def add_columns(self) -> pd.DataFrame:
        '''依據expCol補齊所需的欄位'''
        return self.df.reindex(columns=list(set(self.oc.fin_cols + self.df.columns.to_list())))

    def to_excel(self) -> None:
        self.df.reindex(columns=self.oc.fin_cols).to_excel(self.file_name, index=False)
        logging.info(f'{self.file_name} saved!')

    def run(self):
        # 如有複數來源檔案須將檔案合併
        self.df = self.concat_fr()
        if self.df.empty:
            return None
        logging.info(f'正在轉檔：{'、'.join([f.file for f in self.fr])}')
        # 付款代號
        self.multi_condition(self.pay_code, [self.oc.pay_code])
        # 辨別是否為該訂單第一件商品
        self.df[self.count] = self.df.groupby(self.oc.code).cumcount()
        # 訂單編號
        self.df[self.oc.code] = self.df[self.oc.code].str.replace(r'#', '', regex=True)
        # 郵遞區號取前三碼
        self.df[self.oc.post_code] = self.df[self.oc.post_code].fillna('').apply(lambda x: x[:3])
        # 更改訂單成立日期格式
        self.df[self.oc.date] = self.df[self.oc.date].fillna('').astype(str).apply(lambda x: dt.strptime(x, self.time_fmt).strftime('%Y%m%d'))
        # 付款方式
        self.df[self.oc.pay] = self.df[self.oc.pay].str.replace(r'\([^\(|\)]*\)', '', regex=True)
        # 商品總金額
        self.df[self.oc.price] = self.df[self.price.sum].sum(axis=1)
        # 替換空白電話號碼為'****'
        self.df.loc[self.df[self.oc.cel].isna(), self.oc.cel] = '****'
        # 商品折扣補0
        self.df[self.oc.discount] = self.df[self.oc.discount].fillna(0)
        # 如果在後續金額計算中需要把商品價格*購買數量
        if self.price.col:
            self.df[self.tmp] = self.df[self.price.col] * self.df[self.oc.number]
        # shopline商品總金額要減掉運費
        if self.oc.fr in [ColumnType().shopline]:
            # 刪除shopline已取消\非貨到付款且未付款的訂單
            self.df = self.df.query(f'訂單狀態 != "已取消"')
            self.df = self.df[~((self.df[self.oc.pay_code] != 6) & (self.df[self.oc.pay_code] != 3) & (self.df['付款狀態'] == '未付款'))]
            # shopline商品總金額要減掉運費
            self.df[self.oc.price] = self.df[self.oc.price] - self.df['運費']
        # rakuten商品總金額要減掉優惠券
        if self.oc.fr in [ColumnType().rakuten]:
            self.df[self.oc.price] = self.df[self.oc.price] - self.df[self.oc.discount]
        if self.oc.fr in [ColumnType().shopee]:
            # shopee刪除非手機號碼字串
            self.df[self.oc.cel] = self.df[self.oc.cel].replace(r'#\d$', '', regex=True)
        if self.oc.fr in [ColumnType().shopee, ColumnType().shopline, ColumnType().rakuten]:
            self.df[self.oc.price] = self.df.groupby(self.oc.code)[self.oc.price].transform('first') / self.df.groupby(self.oc.code)[self.tmp].transform(lambda x: x.sum()) * self.df[self.tmp]
        # 金流費用
        self.multi_condition(self.cash_fee, [self.oc.cash_fee])
        if self.oc.fr in [ColumnType().rakuten]:
            # rakuten金流費用
            self.df[self.oc.cash_fee] = self.df[self.oc.price] * self.df[self.oc.cash_fee].fillna(0)
        # 成交手續費
        self.df[self.oc.fee] = self.df[self.oc.price] * self.feeRate.pc + self.feeRate.add
        # 訂單處理費
        self.df[self.oc.order_fee] = (self.df.groupby(self.oc.code)[self.oc.code].transform("count") - 1) * 10 + 26
        # shopline, rakuten運費
        if self.oc.fr in [ColumnType().shopline, ColumnType().rakuten]:
            ship = self.df.loc[self.df['運費'] > 0].copy()
            ship[self.oc.product_code] = '888888888'
            ship[self.oc.price] = ship['運費']
            ship = ship[self.oc.ship_cols]
            self.df = pd.concat([self.df, ship]).sort_values(by=[self.oc.date, self.oc.code])
        self.df.loc[self.df[self.oc.product_code] == '888888888', self.oc.number] = 1
        # 合併商品總表中的資訊並計算撿貨費
        self.df = self.product_detail()
        # 撿貨費
        self.df[self.oc.tally] = self.df['撿貨數量'] * self.df[self.oc.number] * 7.5
        # 進貨小計
        self.df[self.oc.purchase_subtotal] = self.df[self.oc.purchase_price] * self.df[self.oc.number]
        if self.oc.fr in [ColumnType().yahoo]:
            # 金流費用
            self.df.loc[self.df[self.oc.pay_code] == 4, self.oc.cash_fee] = self.df[self.oc.price] * self.df[self.oc.cash_fee]
            # 訂單金額
            self.df[self.oc.subtotal] = self.df.groupby(self.oc.code)[self.price.sum].transform(lambda x: x.sum()).sum(axis=1)
        elif self.oc.fr in [ColumnType().shopee, ColumnType().shopline, ColumnType().rakuten]:
            # 訂單金額
            self.df[self.oc.subtotal] = self.df[self.price.sum].sum(axis=1)
        # rakuten物流交寄使用費
        self.multi_condition(self.delivery_fee, [self.oc.delivery_fee])
        # 運費
        self.multi_condition(self.ship, [self.oc.ship])
        # shopee隱碼服務費
        self.multi_condition(self.service_fee, [self.oc.service_fee])
        # 依倉庫調整撿貨費、訂單處理費、運費
        self.df.loc[self.df[self.oc.warehouse].fillna('').str.contains(r'^(?:原廠出貨|公司倉)$', regex=True), [self.oc.tally, self.oc.order_fee]] = 0
        self.df.loc[self.df[self.oc.warehouse] == '原廠出貨', self.oc.ship] = 0
        # 利潤
        self.df[self.oc.subtotal] = self.df.groupby(self.oc.code)[self.oc.subtotal].transform('first')
        self.df[self.oc.profit] = self.df[self.oc.subtotal] - self.df.groupby(self.oc.code)[self.oc.profit_cols].transform(lambda x: x.sum()).sum(axis=1)
        # 利潤百分比
        self.df[self.oc.profit_pc] = (self.df[self.oc.profit] / self.df[self.oc.profit_denominator] * 100).round(2).astype(str) + '%'
        # 補齊需要的欄位
        self.df = self.add_columns()
        # 如果不是第一件商品，則'訂單金額','利潤','利潤百分比','運費','訂單處理費','隱碼服務費'為0
        self.df.loc[self.df[self.count] != 0, [self.oc.subtotal, self.oc.profit, self.oc.profit_pc, self.oc.ship, self.oc.order_fee, self.oc.service_fee]] = 0
        # 價格四捨五入至整數
        self.round_cols = [col for col in self.oc.round_cols if col in self.df.columns]
        self.df[self.round_cols] = self.df[self.round_cols].fillna(0).round(0)
        # 匯出需要的欄位
        self.to_excel()


def main():
    [os.remove(f'{result}/{f}') for f in get_files_name(result, 'xlsx') if os.path.exists(f'{result}/{f}')]
    yahoo_mall = Converter(
        fr=[SourceFiles().yahoo_mall],
        cov={'交易序號': str, '訂單編號': str, '店家商品料號': str, '收件人電話(日)': str, '收件人行動電話': str, '收件人郵遞區號': str, '轉單日期': str},
        oc=OutputColumns(ColumnType().yahoo),
        price=Price(['金額小計', '超贈點折抵金額', '行銷補助金額']),
        time_fmt='%Y/%m/%d %H:%M',
        file_name='yahoo購物中心'
    )
    shopee = Converter(
        fr=[SourceFiles().shopee],
        cov={'訂單編號': str, '商品選項貨號': str, '收件者電話': str, '取件門市店號': str, '郵遞區號': str, '訂單成立日期': str, '蝦皮專線和包裹查詢碼 \n(請複製下方完整編號提供給您配合的物流商當做聯絡電話)': str},
        oc=OutputColumns(ColumnType().shopee),
        price=Price(['買家總支付金額', '蝦幣折抵', '銀行信用卡活動折抵', '優惠券'], '商品活動價格'),
        time_fmt='%Y-%m-%d %H:%M',
        file_name='shopee'
    )
    shopline = Converter(
        fr=[SourceFiles().shopline],
        cov={'訂單號碼': str, '郵政編號（如適用)': str, '電話號碼': str, '收件人電話號碼': str, '訂單成立日期': str, '商品貨號': str, '全家服務編號 / 7-11 店號': str},
        oc=OutputColumns(ColumnType().shopline),
        price=Price(['付款總金額'], '商品結帳價'),
        time_fmt='%Y-%m-%d %H:%M:%S',
        file_name='shopline'
    )
    rakuten = Converter(
        fr=[SourceFiles().rakuten],
        cov={'訂單日期': str, '訂單號碼': str, '收件人的電話號碼': str, '目的地郵遞區號': str, '商品管理編號 (SKU)': str},
        oc=OutputColumns(ColumnType().rakuten),
        price=Price(['商品總金額'], '商品總金額'),
        time_fmt='%Y-%m-%d %H:%M:%S',
        file_name='rakuten'
    )
    [cov.run() for cov in [shopee, shopline, yahoo_mall, rakuten]]


if __name__ == '__main__':
    logFile = '設定/run.log'
    logging.basicConfig(format='%(asctime)s %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S', level=logging.INFO, handlers=[logging.FileHandler(logFile), logging.StreamHandler()])
    try:
        main()
    except:
        logging.exception(f'錯誤訊息已處存至 {logFile}')
    input('按Enter繼續...')
