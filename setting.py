import os
import logging
import tkinter as tk
from pickle import dump, load

os.chdir(os.path.dirname(os.path.abspath(__file__)))

feeRate = 'feeRate'
rate = 'rate'
add = 'add'
password = 'password'
setting_pkl = '設定/settings.pkl'
settings_default = {
    'yahoo商城': {'feeRate': {'rate': 0.0568, 'add': 2}, 'password': ''},
    'yahoo購物中心': {'feeRate': {'rate': 0.15, 'add': 2}, 'password': ''},
    'shopee': {'feeRate': {'rate': 0.135, 'add': 0}, 'password': ''},
    'shopline': {'feeRate': {'rate': 0.028, 'add': 0}, 'password': ''},
    'rakuten': {'feeRate': {'rate': 0, 'add': 0}, 'password': ''}
}
# dump(settings_default, open(setting_pkl, 'wb'))


class GUI:
    def __init__(self, root):
        # 如果有設定檔則讀取
        self.settings = load(open(setting_pkl, 'rb')) if os.path.exists(setting_pkl) else settings_default
        # 如果更新後sources數量增加則新增預設數值
        self.settings = {k: {**v, **self.settings[k]} for k, v in settings_default.items()}
        self.sources = list(self.settings.keys())
        self.root = root
        self.root.geometry('500x400')
        self.root.title('轉檔設定')
        self.root.resizable(False, False)
        self.root.config(padx=10, pady=10,)
        self.source_files_value = tk.StringVar(value=self.sources[0])
        self.source_prev = self.source_files_value.get()
        self.font = ('微軟正黑體', 12)

        tk.Label(self.root, font=self.font, text='來源檔案:', justify=tk.LEFT).grid(row=0, column=0, sticky=tk.W)
        select_source = tk.OptionMenu(self.root, self.source_files_value, *self.sources)
        select_source.config(font=self.font)
        select_source.grid(row=0, column=1, columnspan=2, pady=3)
        option = root.nametowidget(select_source.menuname)
        option.config(font=self.font)

        self.feeRateRatio = tk.StringVar(value=self.settings[self.source_files_value.get()][feeRate][rate])
        tk.Label(self.root, font=self.font, text='手續費:', justify=tk.LEFT).grid(row=1, column=0, pady=3, sticky=tk.W)
        tk.Label(self.root, font=self.font, text='商品總金額').grid(row=1, column=1, pady=3)
        tk.Label(self.root, font=self.font, text='×').grid(row=1, column=2, pady=3)
        tk.Entry(self.root, font=self.font, textvariable=self.feeRateRatio, width=10).grid(row=1, column=3, pady=3)

        self.feeRateAdd = tk.StringVar(value=self.settings[self.source_files_value.get()][feeRate][add])
        tk.Label(self.root, font=self.font, text='+').grid(row=1, column=4, pady=3)
        tk.Entry(self.root, font=self.font, textvariable=self.feeRateAdd, width=10).grid(row=1, column=5, pady=3)

        self.filePassword = tk.StringVar(value=self.settings[self.source_files_value.get()][password])
        tk.Label(self.root, font=self.font, text='密碼:', justify=tk.LEFT).grid(row=2, column=0, sticky=tk.W)
        tk.Entry(self.root, font=self.font, textvariable=self.filePassword, width=10).grid(row=2, column=1, pady=3)

        tk.Button(self.root, font=self.font, text='儲存並關閉', command=self.save).grid(row=3, column=0, columnspan=6, pady=3)
        self.source_files_value.trace_add('write', self.switch_source)

    def save(self):
        self.settings.update({self.source_files_value.get(): {
            feeRate: {rate: float(self.feeRateRatio.get()), add: float(self.feeRateAdd.get())},
            password: self.filePassword.get()
        }})
        dump(self.settings, open(setting_pkl, 'wb'))
        self.root.destroy()

    def switch_source(self, *e):
        self.settings.update({self.source_prev: {
            feeRate: {rate: float(self.feeRateRatio.get()), add: float(self.feeRateAdd.get())},
            password: self.filePassword.get()
        }})
        self.source_prev = self.source_files_value.get()
        self.feeRateRatio.set(self.settings[self.source_files_value.get()][feeRate][rate])
        self.feeRateAdd.set(self.settings[self.source_files_value.get()][feeRate][add])
        self.filePassword.set(self.settings[self.source_files_value.get()][password])


def main():
    updater = GUI(tk.Tk())
    updater.root.mainloop()


if __name__ == '__main__':
    logFile = '設定/run.log'
    os.makedirs('設定', exist_ok=True)
    logging.basicConfig(format='%(asctime)s %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S', level=logging.INFO, handlers=[logging.FileHandler(logFile), logging.StreamHandler()])
    try:
        main()
    except:
        logging.exception(f'錯誤訊息已處存至 {logFile}')
