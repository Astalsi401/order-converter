#### 計算方式

**付款代號**

| 代號 | 欄位值                                                                                                |
| ---- | ----------------------------------------------------------------------------------------------------- |
| 1    | 付款方式=['銀行轉帳', '蝦皮錢包', '線上支付', 'ATM/銀行轉帳', 'ATM', '全家繳費', 'ATM 轉帳']          |
| 3    | 付款方式=['貨到付款', '現付', '7-11 門市取貨付款']                                                    |
| 4    | 付款方式=['信用卡', '信用卡分期付款', 'LINE Pay', '信用卡付款', '信用卡一次', '分期付款', '街口支付'] |
| 6    | 付款方式=['7-11', '7-11 門市取貨付款', '全家門市取貨付款']                                            |
| 6    | 付款方式=['貨到付款'] **_且_** 配送方式=['7-ELEVEN','7-11 取貨 (到店付款)', '全家取貨 (到店付款)']    |

**金流費用**

| 金額 | 欄位值                                                                                    |
| ---- | ----------------------------------------------------------------------------------------- |
| 0    | 付款方式=['銀行轉帳', '蝦皮錢包', '線上支付', 'ATM/銀行轉帳', 'ATM', '全家繳費']          |
| 2%   | 付款方式=['信用卡', '信用卡分期付款', '信用卡付款', '信用卡一次', '分期付款', '街口支付'] |
| 2.2% | 付款方式=['LINE Pay']                                                                     |
| 15   | 付款方式=['全家繳費']                                                                     |
| 48   | 付款方式=['7-11']                                                                         |

**物流交寄使用費**

| 金額 | 欄位值                                     |
| ---- | ------------------------------------------ |
| 48   | 配送方式=['全家門市取貨', '7-11 門市取貨'] |

**運費**

| 金額 | 欄位值                                                                                                |
| ---- | ----------------------------------------------------------------------------------------------------- |
| 75   | 配送方式=['7-ELEVEN', '7-11 取貨 (到店付款)', '全家取貨 (到店付款)', '全家門市取貨', '7-11 門市取貨'] |
| 140  | 外站名稱=['yahoo 購物中心']                                                                           |
| 140  | 配送方式=['賣家宅配', '宅配', '常溫宅配(倉儲中心)']                                                   |

**隱碼服務費**

| 金額 | 欄位值                                                               |
| ---- | -------------------------------------------------------------------- |
| 0    | 配送方式=['7-ELEVEN', '7-11 取貨 (到店付款)', '全家取貨 (到店付款)'] |
| 10   | 配送方式=['賣家宅配', '宅配']                                        |

#### 計算流程

- sum() 代表整個訂單所有商品加總

**shopline 計算流程**

1. 建立'付款代號'
1. '商品總金額'='付款總金額'
1. '電話號碼'為空白則補'\*\*\*\*'
1. '商品折扣'如為空白則補 0
1. 郵遞區號取前 3 碼
1. 更改訂單成立日期格式：%Y%m%d
1. 刪除訂單狀態=已取消
1. '商品總金額'='商品總金額'-'運費(箱子+包材+運費)'
1. 根據當前'商品總金額'計算'金流費用'
1. 根據當前'商品總金額'計算'成交手續費'
1. '訂單處理費'=26，每筆訂單一次
1. '商品總金額'=加總訂單總金額 / sum(各商品價格 × 購買數量) × (各商品價格 × 購買數量)
1. 將運費獨立列出
1. 合併商品總表欄位
1. '入庫出貨撿貨費'='撿貨數量'×'購買數量'×7.5
1. '進貨小計'='進貨價'×'購買數量'
1. '訂單金額'='付款總金額'
1. 計算'運費(箱子+包材+運費)'
1. 依倉庫調整'入庫出貨撿貨費'、'訂單處理費'，如果原廠出貨、公司倉，'訂單處理費'=0, '入庫出貨撿貨費'=0
1. 依倉庫調整'運費(箱子+包材+運費)'，如果原廠出貨，'運費(箱子+包材+運費)'=0
1. 如果不是第一件商品，則'訂單金額','運費','訂單處理費','隱碼服務費','點數成本負擔'為 0
1. '利潤'=sum('訂單金額')-sum('進貨小計'+'成交手續費(含購物車費用)'+'入庫出貨撿貨費'+'訂單處理費'+'運費(箱子+包材+運費)')
1. '利潤百分比'='利潤'/'訂單金額'×100
1. 如果不是第一件商品，則'利潤','利潤百分比'為 0
1. 空白欄位補 0 並四捨五入至小數點後 1 位：['商品總金額', '進貨價', '進貨小計', '成交手續費', '金流費用', '物流交寄使用費', '點數成本負擔', '訂單金額']

**rakuten 計算流程**

1. 建立'付款代號'
1. '商品總金額'='訂單與運費總和'
1. '電話號碼'為空白則補'\*\*\*\*'
1. '商品折扣'如為空白則補 0
1. 郵遞區號取前 3 碼
1. 更改訂單成立日期格式：%Y%m%d
1. '商品總金額'='商品總金額'-'運費(箱子+包材+運費)'-'優惠券'
1. 除非是訂單中第一筆商品，否則商品總金額、點數成本負擔應為 0
1. 根據當前'商品總金額'計算'成交手續費'
1. 根據當前'商品總金額'計算'金流費用'
1. '訂單處理費'=26，每筆訂單一次
1. '商品總金額'=加總訂單總金額 / sum(各商品價格 × 購買數量) × (各商品價格 × 購買數量)
1. 將運費獨立列出
1. 合併商品總表欄位
1. '入庫出貨撿貨費'='撿貨數量'×'購買數量'×7.5
1. '進貨小計'='進貨價'×'購買數量'
1. '訂單金額'=sum('商品總金額')
1. 計算'物流交寄使用費'
1. 計算'運費(箱子+包材+運費)'
1. 依倉庫調整'入庫出貨撿貨費'、'訂單處理費'，如果原廠出貨、公司倉，'訂單處理費'=0, '入庫出貨撿貨費'=0
1. 依倉庫調整'運費(箱子+包材+運費)'，如果原廠出貨，'運費(箱子+包材+運費)'=0
1. 如果不是第一件商品，則'訂單金額','運費','訂單處理費','隱碼服務費','點數成本負擔'為 0
1. '利潤'=sum('訂單金額')-sum('進貨小計'+'成交手續費(含購物車費用)'+'金流費用'+'物流交寄使用費'+'點數成本負擔'+'入庫出貨撿貨費'+'訂單處理費'+'運費(箱子+包材+運費)')
1. '利潤百分比'='利潤'/'訂單金額'×100
1. 如果不是第一件商品，則'利潤','利潤百分比'為 0
1. 空白欄位補 0 並四捨五入至小數點後 1 位：['商品總金額', '進貨價', '進貨小計', '成交手續費', '金流費用', '物流交寄使用費', '點數成本負擔', '訂單金額']

**shopee 計算流程**

1. 建立'付款代號'
1. '商品總金額'='買家總支付金額'+'蝦幣折抵'+'銀行信用卡活動折抵'+'優惠券'
1. '電話號碼'為空白則補'\*\*\*\*'
1. '商品折扣'如為空白則補 0
1. 郵遞區號取前 3 碼
1. 更改訂單成立日期格式：%Y%m%d
1. 刪除手機號碼尾段的#+數字
1. 根據當前'商品總金額'計算'金流費用'
1. 根據當前'商品總金額'計算'成交手續費'
1. '訂單處理費'=26，每筆訂單一次
1. '商品總金額'=加總訂單總金額 / sum(各商品價格 × 購買數量) × (各商品價格 × 購買數量)
1. 合併商品總表欄位
1. '入庫出貨撿貨費'='撿貨數量'×'購買數量'×7.5
1. '進貨小計'='進貨價'×'購買數量'
1. '訂單金額'='買家總支付金額'+'蝦幣折抵'+'銀行信用卡活動折抵'+'優惠券'
1. 計算'運費(箱子+包材+運費)'
1. 計算'隱碼服務費'
1. 依倉庫調整'入庫出貨撿貨費'、'訂單處理費'，如果原廠出貨、公司倉，'訂單處理費'=0, '入庫出貨撿貨費'=0
1. 依倉庫調整'運費(箱子+包材+運費)'，如果原廠出貨，'運費(箱子+包材+運費)'=0
1. 如果不是第一件商品，則'訂單金額','運費','訂單處理費','隱碼服務費','點數成本負擔'為 0
1. '利潤'=sum('訂單金額')-sum('進貨小計'+'成交手續費(含購物車費用)'+'蝦幣回饋券'+'入庫出貨撿貨費'+'訂單處理費'+'運費(箱子+包材+運費)'+'隱碼服務費')
1. '利潤百分比'='利潤'/'訂單金額'×100
1. 如果不是第一件商品，則'利潤','利潤百分比'為 0
1. 空白欄位補 0 並四捨五入至小數點後 1 位：['商品總金額', '進貨價', '進貨小計', '成交手續費', '金流費用', '物流交寄使用費', '點數成本負擔', '訂單金額']

**yahoo 購物中心**

1. 建立'付款代號'
1. 郵遞區號取前 3 碼
1. 更改訂單成立日期格式：%Y%m%d
1. '商品總金額'='金額小計'+'超贈點折抵金額'+'行銷補助金額'
1. '電話號碼'為空白則補'\*\*\*\*'
1. '商品折扣'如為空白則補 0
1. 根據當前'商品總金額'計算'金流費用'
1. 根據當前'商品總金額'計算'成交手續費'
1. '訂單處理費'=26，每筆訂單一次
1. 合併商品總表欄位
1. '入庫出貨撿貨費'='撿貨數量'×'購買數量'×7.5
1. '進貨小計'='進貨價'×'購買數量'
1. '訂單金額'=sum('金額小計'+'超贈點折抵金額'+'行銷補助金額')
1. 計算'運費(箱子+包材+運費)'
1. 依倉庫調整'入庫出貨撿貨費'、'訂單處理費'，如果原廠出貨、公司倉，'訂單處理費'=0, '入庫出貨撿貨費'=0
1. 依倉庫調整'運費(箱子+包材+運費)'，如果原廠出貨，'運費(箱子+包材+運費)'=0
1. 如果不是第一件商品，則'訂單金額','運費','訂單處理費','隱碼服務費','點數成本負擔'為 0
1. '利潤'=sum('訂單金額')-sum('進貨小計'+'成交手續費(含購物車費用)'+'金流費用'+'入庫出貨撿貨費'+'訂單處理費'+'運費(箱子+包材+運費)')
1. '利潤百分比'='利潤'/'訂單金額'×100
1. 如果不是第一件商品，則'利潤','利潤百分比'為 0
1. 空白欄位補 0 並四捨五入至小數點後 1 位：['商品總金額', '進貨價', '進貨小計', '成交手續費', '金流費用', '物流交寄使用費', '點數成本負擔', '訂單金額']
