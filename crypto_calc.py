import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

class CryptoCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("暗号資産損益計算機")
        self.root.geometry("1400x900")
        
        # データフレームを保存する変数
        self.df = None
        self.result_df = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # タイトル
        title_label = ttk.Label(main_frame, text="暗号資産損益計算機", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 20))
        
        # ファイル選択部分
        file_frame = ttk.LabelFrame(main_frame, text="ファイル選択", padding="10")
        file_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=80)
        file_entry.grid(row=0, column=0, padx=(0, 10))
        
        browse_btn = ttk.Button(file_frame, text="ファイル選択", command=self.browse_file)
        browse_btn.grid(row=0, column=1)
        
        load_btn = ttk.Button(file_frame, text="データ読み込み", command=self.load_data)
        load_btn.grid(row=0, column=2, padx=(10, 0))
        
        # アルゴリズム説明部分
        algo_frame = ttk.LabelFrame(main_frame, text="計算アルゴリズム", padding="10")
        algo_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 20))
        
        algo_text = tk.Text(algo_frame, height=6, width=120, wrap=tk.WORD)
        algo_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # アルゴリズムの説明を挿入
        algorithm_description = """
【移動平均法による損益計算アルゴリズム】

1. 買い取引（BUY）の場合：
   - 保有数量に取引数量を加算
   - 累計取得費に取引金額と手数料を加算
   - 平均取得単価 = 累計取得費 ÷ 保有数量

2. 売り取引（SELL）の場合：
   - 現在の平均取得単価で取得費を計算
   - 損益 = 売却代金 - 手数料 - 取得費
   - 保有数量と累計取得費から売却分を減算

3. 年間損益合計：
   - 全売却取引の損益を合計

※ この方法は税務署が認める移動平均法に基づいています

【計算対象となる取引】
• Transaction Buy: 暗号資産購入
• Transaction Sold: 暗号資産売却
• Transaction Spend: 暗号資産支払い
• Transaction Revenue: 暗号資産受取
• Transaction Fee: 手数料
• Simple Earn Locked Rewards: 简单赚币锁定奖励

【計算対象外となる操作】
• Simple Earn Flexible Rewards: 简单赚币灵活奖励
• Deposit/Withdraw: 充值/提现
• Transfer In/Out: 转入/转出

【ETH/BTC历史价格功能】
• 点击"自动获取ETH/BTC历史价格"按钮自动从网络获取历史价格
• 使用CoinGecko API免费获取ETH和BTC的日元历史价格
• 历史价格自动保存到historical_prices.json文件中
• "损益计算执行"按钮会优先使用历史价格，避免估值误差
• 如果没有历史价格，会回退到原来的计算方法
        """
        algo_text.insert(tk.END, algorithm_description)
        algo_text.config(state=tk.DISABLED)
        
        # 計算実行ボタン
        calc_btn = ttk.Button(main_frame, text="損益計算実行", command=self.calculate_profit)
        calc_btn.grid(row=3, column=0, pady=20)
        
        # 获取历史价格按钮
        get_price_btn = ttk.Button(main_frame, text="自动获取ETH/BTC历史价格", command=self.get_historical_prices)
        get_price_btn.grid(row=3, column=1, pady=20)
        
        # 只计算日元交易按钮
        calc_jpy_btn = ttk.Button(main_frame, text="日元交易のみ計算", command=self.calculate_jpy_only)
        calc_jpy_btn.grid(row=3, column=2, pady=20)
        
        # 左側：公式表示部分
        formula_frame = ttk.LabelFrame(main_frame, text="計算公式詳細", padding="10")
        formula_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10), pady=(0, 20))
        
        self.formula_text = tk.Text(formula_frame, height=20, width=70, wrap=tk.WORD)
        formula_scrollbar = ttk.Scrollbar(formula_frame, orient=tk.VERTICAL, command=self.formula_text.yview)
        self.formula_text.configure(yscrollcommand=formula_scrollbar.set)
        
        self.formula_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        formula_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 初期公式説明
        initial_formula = "計算を実行すると、ここに各取引の詳細な計算公式が表示されます。"
        self.formula_text.insert(tk.END, initial_formula)
        self.formula_text.config(state=tk.DISABLED)
        
        # 右側：結果表示部分
        result_frame = ttk.LabelFrame(main_frame, text="計算結果", padding="10")
        result_frame.grid(row=4, column=2, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0), pady=(0, 20))
        
        # 結果表示用のTreeview
        self.tree = ttk.Treeview(result_frame, columns=("番号", "日付", "種別", "通貨", "数量", "単価(JPY)", "損益(JPY)", "平均取得単価"), show="headings", height=18)
        
        # カラムの設定
        self.tree.heading("番号", text="番号")
        self.tree.heading("日付", text="日付")
        self.tree.heading("種別", text="種別")
        self.tree.heading("通貨", text="通貨")
        self.tree.heading("数量", text="数量")
        self.tree.heading("単価(JPY)", text="単価(JPY)")
        self.tree.heading("損益(JPY)", text="損益(JPY)")
        self.tree.heading("平均取得単価", text="平均取得単価")
        
        # カラム幅の設定
        self.tree.column("番号", width=50)
        self.tree.column("日付", width=100)
        self.tree.column("種別", width=80)
        self.tree.column("通貨", width=80)
        self.tree.column("数量", width=100)
        self.tree.column("単価(JPY)", width=120)
        self.tree.column("損益(JPY)", width=120)
        self.tree.column("平均取得単価", width=120)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 年間損益合計表示（下部中央）
        total_frame = ttk.Frame(main_frame)
        total_frame.grid(row=5, column=0, columnspan=4, pady=(0, 20))
        
        self.total_profit_var = tk.StringVar()
        self.total_profit_var.set("年間損益合計: 計算してください")
        total_profit_label = ttk.Label(total_frame, textvariable=self.total_profit_var, font=("Arial", 14, "bold"))
        total_profit_label.pack()
        
        # グリッドの重み設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.columnconfigure(3, weight=1)
        main_frame.rowconfigure(4, weight=1)
        formula_frame.columnconfigure(0, weight=1)
        formula_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
            
    def load_data(self):
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("エラー", "ファイルを選択してください")
            return
            
        try:
            # 文字コードを自動判定
            encodings = ['utf-8', 'shift-jis', 'cp932']
            for encoding in encodings:
                try:
                    self.df = pd.read_csv(file_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                messagebox.showerror("エラー", "ファイルの文字コードが判別できません")
                return
                
            messagebox.showinfo("成功", f"データを読み込みました\n行数: {len(self.df)}\n列数: {len(self.df.columns)}")
            
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました:\n{str(e)}")
            
    def calculate_profit(self):
        if self.df is None:
            messagebox.showerror("エラー", "先にデータを読み込んでください")
            return
            
        try:
            # 取引データのみを抽出（入出金、報酬などは除外）
            trade_operations = [
                'Transaction Buy',           # 暗号資産購入
                'Transaction Sold',          # 暗号資産売却
                'Transaction Spend',         # 暗号資産支払い
                'Transaction Revenue',       # 暗号資産受取
                'Transaction Fee',           # 手数料
                'Simple Earn Locked Rewards' # 简单赚币锁定奖励
            ]
            
            # 取引データのみをフィルタリング
            trade_df = self.df[self.df['Operation'].isin(trade_operations)].copy()
            
            if len(trade_df) == 0:
                messagebox.showwarning("警告", "計算可能な取引データが見つかりませんでした")
                return
            
            # 日付でソート
            trade_df['UTC_Time'] = pd.to_datetime(trade_df['UTC_Time'])
            trade_df = trade_df.sort_values('UTC_Time')
            
            # デバッグ用：取引データの概要を表示
            print(f"取引データ総数: {len(trade_df)}")
            print("取引タイプ別件数:")
            print(trade_df['Operation'].value_counts())
            
            # 計算用の変数
            holdings = {}  # 通貨ごとの保有数量と取得費
            results = []
            formula_details = []  # 公式詳細を保存
            processed_transactions = set()  # 処理済み取引を記録（重複防止）
            transaction_counter = 0  # 取引番号カウンター
            
            def format_number(value, decimal_places=8):
                """数字を適切な形式でフォーマット"""
                if abs(value) < 0.000001:  # 1e-06未満
                    return f"{value:.8f}"  # 8桁の小数で表示
                elif abs(value) < 0.01:    # 1銭未満
                    return f"{value:.6f}"  # 6桁の小数で表示
                elif abs(value) < 1:       # 1未満
                    return f"{value:.4f}"  # 4桁の小数で表示
                else:
                    return f"{value:.2f}"  # 2桁の小数で表示

            # 取引ペアを特定するためのヘルパー関数
            def find_transaction_pair(timestamp, operation, coin):
                """特定のタイムスタンプで関連する取引を検索"""
                related_transactions = []
                
                # 同じタイムスタンプの取引を検索
                same_time_transactions = trade_df[trade_df['UTC_Time'] == timestamp]
                
                if operation == 'Transaction Buy':
                    # 購入の場合：支払い通貨を探す
                    if coin == 'ETH':
                        # ETH購入の場合、BTC支払いを探す
                        btc_spend = same_time_transactions[
                            (same_time_transactions['Operation'] == 'Transaction Spend') & 
                            (same_time_transactions['Coin'] == 'BTC')
                        ]
                        if not btc_spend.empty:
                            related_transactions.extend(btc_spend.to_dict('records'))
                    elif coin == 'BTC':
                        # BTC購入の場合、日元支払いを探す
                        jpy_spend = same_time_transactions[
                            (same_time_transactions['Operation'] == 'Transaction Spend') & 
                            (same_time_transactions['Coin'] == 'JPY')
                        ]
                        if not jpy_spend.empty:
                            related_transactions.extend(jpy_spend.to_dict('records'))
                    
                    # 購入取引自体を追加
                    buy_transaction = same_time_transactions[
                        (same_time_transactions['Operation'] == 'Transaction Buy') & 
                        (same_time_transactions['Coin'] == coin)
                    ]
                    if not buy_transaction.empty:
                        related_transactions.extend(buy_transaction.to_dict('records'))
                
                elif operation == 'Transaction Sold':
                    # 売却の場合：受取通貨を探す
                    if coin == 'ETH':
                        # ETH売却の場合、BTC受取を探す
                        btc_revenue = same_time_transactions[
                            (same_time_transactions['Operation'] == 'Transaction Revenue') & 
                            (same_time_transactions['Coin'] == 'BTC')
                        ]
                        if not btc_revenue.empty:
                            related_transactions.extend(btc_revenue.to_dict('records'))
                    elif coin == 'BTC':
                        # BTC売却の場合、日元受取を探す
                        jpy_revenue = same_time_transactions[
                            (same_time_transactions['Operation'] == 'Transaction Revenue') & 
                            (same_time_transactions['Coin'] == 'JPY')
                        ]
                        if not jpy_revenue.empty:
                            related_transactions.extend(jpy_revenue.to_dict('records'))
                    
                    # 売却取引自体を追加
                    sell_transaction = same_time_transactions[
                        (same_time_transactions['Operation'] == 'Transaction Sold') & 
                        (same_time_transactions['Coin'] == coin)
                    ]
                    if not sell_transaction.empty:
                        related_transactions.extend(sell_transaction.to_dict('records'))
                
                # 手数料を追加
                fees = same_time_transactions[
                    (same_time_transactions['Operation'] == 'Transaction Fee') & 
                    (same_time_transactions['Coin'] == coin)
                ]
                if not fees.empty:
                    related_transactions.extend(fees.to_dict('records'))
                
                return related_transactions

            # ETH/BTC取引の価格計算用のヘルパー関数
            def calculate_eth_btc_price(eth_quantity, btc_quantity):
                """ETH/BTC取引の価格を計算"""
                if eth_quantity > 0 and btc_quantity > 0:
                    return btc_quantity / eth_quantity
                return 0

            # BTCの平均取得単価を取得する関数
            def get_btc_avg_cost():
                """BTCの平均取得単価を取得"""
                if 'BTC' in holdings and holdings['BTC']['quantity'] > 0:
                    avg_cost = holdings['BTC']['total_cost'] / holdings['BTC']['quantity']
                    # BTCの価格が異常に高い場合は、合理的な価格に制限
                    if avg_cost > 10000000:  # 1000万円以上の場合
                        print(f"  ⚠️ BTC単価が異常に高いため、計算をスキップ: {avg_cost:,.0f} 円")
                        return 0
                    return avg_cost
                return 0

            for idx, row in trade_df.iterrows():
                operation = row['Operation']
                coin = row['Coin']
                change = float(row['Change'])
                date = row['UTC_Time'].strftime('%Y-%m-%d %H:%M')
                
                # 处理所有加密货币交易
                # if coin != 'ETH':
                #     continue
                
                # 重複処理防止
                transaction_key = f"{date}_{operation}_{coin}_{change}"
                if transaction_key in processed_transactions:
                    print(f"重複取引をスキップ: {transaction_key}")
                    continue
                
                # デバッグ用：各取引の詳細を表示
                print(f"処理中: {date} | {operation} | {coin} | {change}")
                
                if coin not in holdings:
                    holdings[coin] = {'quantity': 0.0, 'total_cost': 0.0}
                
                if operation == 'Transaction Buy':
                    # 暗号資産購入
                    if change > 0:  # 正の値の場合のみ処理
                        transaction_counter += 1  # 取引番号を増加
                        
                        # 暗号資産の数量
                        quantity = change
                        old_quantity = holdings[coin]['quantity']
                        old_cost = holdings[coin]['total_cost']
                        
                        # 関連する取引を検索（BTC支払いなど）
                        related_transactions = find_transaction_pair(row['UTC_Time'], operation, coin)
                        
                        # 支払い金額を計算
                        jpy_spend_amount = 0.0
                        jpy_fee_amount = 0.0
                        
                        if coin == 'ETH':
                            # ETH購入の場合：まず日元支払いを探す
                            jpy_spend_amount = 0.0
                            btc_spend_amount = 0.0
                            fee_amount = 0.0
                            
                            for related in related_transactions:
                                if related['Operation'] == 'Transaction Spend' and related['Coin'] == 'JPY':
                                    jpy_spend_amount += abs(float(related['Change']))
                                elif related['Operation'] == 'Transaction Spend' and related['Coin'] == 'BTC':
                                    btc_spend_amount += abs(float(related['Change']))
                            
                            # 日元支払いがない場合、BTC支払いを計算
                            if jpy_spend_amount == 0 and btc_spend_amount > 0:
                                # 首先尝试使用历史价格
                                transaction_date = row['UTC_Time'].strftime('%Y-%m-%d')
                                btc_price = self.get_price_for_date(transaction_date, 'BTC')
                                
                                if btc_price and btc_price > 0:
                                    jpy_spend_amount = btc_spend_amount * btc_price
                                    print(f"  → ETH購入（BTC支払い）: {quantity} ETH = {btc_spend_amount:.8f} BTC = {jpy_spend_amount:,.0f} 円 (历史BTC単価: {btc_price:,.0f} 円)")
                                else:
                                    # 如果没有历史价格，使用BTC的平均取得单价
                                    btc_avg_cost = get_btc_avg_cost()
                                    if btc_avg_cost > 0:
                                        jpy_spend_amount = btc_spend_amount * btc_avg_cost
                                        print(f"  → ETH購入（BTC支払い）: {quantity} ETH = {btc_spend_amount:.8f} BTC = {jpy_spend_amount:,.0f} 円 (BTC平均取得単価: {btc_avg_cost:,.0f} 円)")
                                    else:
                                        print(f"  ⚠️ 仮想通貨間取引: {quantity} ETH = {btc_spend_amount:.8f} BTC (BTC取得単価なし)")
                                        continue
                            
                            # 手数料を計算（日元優先）
                            jpy_fee_amount = 0.0
                            for related in related_transactions:
                                if related['Operation'] == 'Transaction Fee' and related['Coin'] == 'JPY':
                                    jpy_fee_amount += abs(float(related['Change']))
                                elif related['Operation'] == 'Transaction Fee' and related['Coin'] == coin:
                                    # 仮想通貨手数料は記録のみ
                                    fee_amount = abs(float(related['Change']))
                                    print(f"  ⚠️ 仮想通貨手数料: {fee_amount:.8f} {coin}")
                            
                        elif coin == 'BTC':
                            # BTC購入の場合：日元支払いを計算
                            for related in related_transactions:
                                if related['Operation'] == 'Transaction Spend' and related['Coin'] == 'JPY':
                                    jpy_spend_amount += abs(float(related['Change']))
                            
                            # 手数料を計算（日元）
                            for related in related_transactions:
                                if related['Operation'] == 'Transaction Fee' and related['Coin'] == 'JPY':
                                    jpy_fee_amount += abs(float(related['Change']))
                        
                        # 支払いがある場合のみ処理
                        if jpy_spend_amount > 0:
                            # 保有数量と取得費を更新（日元）
                            holdings[coin]['quantity'] += quantity
                            holdings[coin]['total_cost'] += jpy_spend_amount + jpy_fee_amount
                            
                            # 公式詳細を記録
                            new_quantity = holdings[coin]['quantity']
                            new_cost = holdings[coin]['total_cost']
                            avg_cost = new_cost / new_quantity if new_quantity > 0 else 0
                            
                            if coin == 'ETH':
                                formula_detail = f"""
【取引{transaction_counter}】{date} {coin}購入取引
購入数量: {format_number(quantity)}
購入金額（BTC）: {btc_spend_amount:.8f} BTC
購入金額（日元）: {jpy_spend_amount:,.0f} 円
手数料: {fee_amount:.8f} {coin}
手数料（日元）: {jpy_fee_amount:,.0f} 円
購入前保有数量: {format_number(old_quantity)}
購入前累計取得費: {old_cost:,.0f} 円
購入後保有数量: {format_number(new_quantity)}
購入後累計取得費: {new_cost:,.0f} 円
平均取得単価: {new_cost:,.0f} 円 ÷ {format_number(new_quantity)} = {avg_cost:,.0f} 円
"""
                            else:
                                formula_detail = f"""
【取引{transaction_counter}】{date} {coin}購入取引
購入数量: {format_number(quantity)}
購入金額（日元）: {jpy_spend_amount:,.0f} 円
手数料: {jpy_fee_amount:,.0f} 円
購入前保有数量: {format_number(old_quantity)}
購入前累計取得費: {old_cost:,.0f} 円
購入後保有数量: {format_number(new_quantity)}
購入後累計取得費: {new_cost:,.0f} 円
平均取得単価: {new_cost:,.0f} 円 ÷ {format_number(new_quantity)} = {avg_cost:,.0f} 円
"""
                            formula_details.append(formula_detail)
                            
                            if coin == 'ETH':
                                print(f"  → {coin}購入: 数量={quantity}, 金額={btc_spend_amount:.8f} BTC ({jpy_spend_amount:,.0f} 円), 手数料={fee_amount:.8f} {coin} ({jpy_fee_amount:,.0f} 円), 累計取得費={holdings[coin]['total_cost']:,.0f} 円")
                            else:
                                print(f"  → {coin}購入: 数量={quantity}, 金額={jpy_spend_amount:,.0f} 円, 手数料={jpy_fee_amount:,.0f} 円, 累計取得費={holdings[coin]['total_cost']:,.0f} 円")
                            
                            results.append([transaction_counter, date, "BUY", coin, quantity, f"{jpy_spend_amount:,.0f}", 0, f"{avg_cost:,.0f}"])
                            
                            # 処理済みとして記録
                            processed_transactions.add(transaction_key)
                            for related in related_transactions:
                                related_key = f"{date}_{related['Operation']}_{related['Coin']}_{related['Change']}"
                                processed_transactions.add(related_key)
                
                elif operation == 'Transaction Sold':
                    # 暗号資産売却
                    if change < 0:  # 負の値の場合のみ処理
                        transaction_counter += 1  # 取引番号を増加
                        
                        quantity = abs(change)
                        formatted_quantity = format_number(quantity)
                        
                        if holdings[coin]['quantity'] >= quantity:
                            # 平均取得単価で損益計算
                            old_quantity = holdings[coin]['quantity']
                            old_cost = holdings[coin]['total_cost']
                            avg_cost = old_cost / old_quantity if old_quantity > 0 else 0
                            cost_basis = avg_cost * quantity
                            
                            # 関連する取引を検索（BTC受取など）
                            related_transactions = find_transaction_pair(row['UTC_Time'], operation, coin)
                            
                            # 売却代金を計算
                            jpy_revenue_amount = 0.0
                            jpy_fee_amount = 0.0
                            
                            if coin == 'ETH':
                                # ETH売却の場合：まず日元受取を探す
                                btc_revenue_amount = 0.0
                                fee_amount = 0.0
                                
                                for related in related_transactions:
                                    if related['Operation'] == 'Transaction Revenue' and related['Coin'] == 'JPY':
                                        jpy_revenue_amount += float(related['Change'])
                                    elif related['Operation'] == 'Transaction Revenue' and related['Coin'] == 'BTC':
                                        btc_revenue_amount += float(related['Change'])
                                
                                # 日元受取がない場合、BTC受取を計算
                                if jpy_revenue_amount == 0 and btc_revenue_amount > 0:
                                    # 首先尝试使用历史价格
                                    transaction_date = row['UTC_Time'].strftime('%Y-%m-%d')
                                    btc_price = self.get_price_for_date(transaction_date, 'BTC')
                                    
                                    if btc_price and btc_price > 0:
                                        jpy_revenue_amount = btc_revenue_amount * btc_price
                                        print(f"  → ETH売却（BTC受取）: {quantity} ETH = {btc_revenue_amount:.8f} BTC = {jpy_revenue_amount:,.0f} 円 (历史BTC単価: {btc_price:,.0f} 円)")
                                    else:
                                        # 如果没有历史价格，使用BTC的平均取得单价
                                        btc_avg_cost = get_btc_avg_cost()
                                        if btc_avg_cost > 0:
                                            jpy_revenue_amount = btc_revenue_amount * btc_avg_cost
                                            print(f"  → ETH売却（BTC受取）: {quantity} ETH = {btc_revenue_amount:.8f} BTC = {jpy_revenue_amount:,.0f} 円 (BTC平均取得単価: {btc_avg_cost:,.0f} 円)")
                                        else:
                                            print(f"  ⚠️ 仮想通貨間取引: {quantity} ETH = {btc_revenue_amount:.8f} BTC (BTC取得単価なし)")
                                            continue
                                
                                # 手数料を計算（日元優先）
                                for related in related_transactions:
                                    if related['Operation'] == 'Transaction Fee' and related['Coin'] == 'JPY':
                                        jpy_fee_amount += abs(float(related['Change']))
                                    elif related['Operation'] == 'Transaction Fee' and related['Coin'] == coin:
                                        # 仮想通貨手数料は記録のみ
                                        fee_amount = abs(float(related['Change']))
                                        print(f"  ⚠️ 仮想通貨手数料: {fee_amount:.8f} {coin}")
                                
                            elif coin == 'BTC':
                                # BTC売却の場合：日元受取を計算
                                for related in related_transactions:
                                    if related['Operation'] == 'Transaction Revenue' and related['Coin'] == 'JPY':
                                        jpy_revenue_amount += float(related['Change'])
                                
                                # 手数料を計算（日元）
                                for related in related_transactions:
                                    if related['Operation'] == 'Transaction Fee' and related['Coin'] == 'JPY':
                                        jpy_fee_amount += abs(float(related['Change']))
                            
                            # 売却代金がある場合のみ処理
                            if jpy_revenue_amount > 0:
                                # 損益計算（手数料を考慮）
                                profit = jpy_revenue_amount - cost_basis - jpy_fee_amount
                                
                                # 公式詳細を記録
                                new_quantity = old_quantity - quantity
                                new_cost = old_cost - cost_basis
                                
                                formula_detail = f"""
【取引{transaction_counter}】{date} {coin}売却取引
売却数量: {formatted_quantity}
売却代金（日元）: {jpy_revenue_amount:,.0f} 円
手数料: {jpy_fee_amount:,.0f} 円
売却前保有数量: {format_number(old_quantity)}
売却前累計取得費: {old_cost:,.0f} 円
平均取得単価: {old_cost:,.0f} 円 ÷ {format_number(old_quantity)} = {avg_cost:,.0f} 円
取得費（売却分）: {avg_cost:,.0f} 円 × {format_number(quantity)} = {cost_basis:,.0f} 円
損益計算: {jpy_revenue_amount:,.0f} 円 - {cost_basis:,.0f} 円 - {jpy_fee_amount:,.0f} 円 = {profit:,.0f} 円
売却後保有数量: {format_number(new_quantity)}
売却後累計取得費: {new_cost:,.0f} 円
"""
                                formula_details.append(formula_detail)
                                
                                print(f"  → {coin}売却: 数量={quantity}, 売却代金={jpy_revenue_amount:,.0f} 円, 取得費={cost_basis:,.0f} 円, 手数料={jpy_fee_amount:,.0f} 円, 損益={profit:,.0f} 円")
                                
                                # 保有数量と取得費を減らす
                                holdings[coin]['quantity'] -= quantity
                                holdings[coin]['total_cost'] -= cost_basis
                                
                                results.append([transaction_counter, date, "SELL", coin, quantity, "N/A", f"{profit:,.0f}", f"{avg_cost:,.0f}"])
                                
                                # 処理済みとして記録
                                processed_transactions.add(transaction_key)
                                for related in related_transactions:
                                    related_key = f"{date}_{related['Operation']}_{related['Coin']}_{related['Change']}"
                                    processed_transactions.add(related_key)
                        else:
                            print(f"警告: {coin}の保有数量が不足しています。必要: {quantity}, 保有: {holdings[coin]['quantity']}")
                
                elif operation == 'Simple Earn Locked Rewards':
                    # 简单赚币锁定奖励
                    if change > 0:  # 正の値の場合のみ処理
                        transaction_counter += 1  # 取引番号を増加
                        
                        # 奖励数量
                        reward_quantity = change
                        
                        # 使用历史价格计算奖励价值
                        transaction_date = row['UTC_Time'].strftime('%Y-%m-%d')
                        coin_price = self.get_price_for_date(transaction_date, coin)
                        
                        if coin_price and coin_price > 0:
                            reward_value = reward_quantity * coin_price
                            print(f"  → {coin}锁定奖励: 数量={reward_quantity}, 历史价格={coin_price:,.0f}円, 价值={reward_value:,.0f}円")
                        else:
                            # 如果没有历史价格，使用默认价格
                            if coin == 'ETH':
                                reward_value = reward_quantity * 17.5  # 1 ETH = 17.5日元
                            elif coin == 'BTC':
                                reward_value = reward_quantity * 50  # 1 BTC = 50日元
                            else:
                                reward_value = 0  # 其他货币暂时设为0
                            print(f"  → {coin}锁定奖励: 数量={reward_quantity}, 默认价格, 价值={reward_value:,.0f}円")
                        
                        # 公式詳細を記録
                        if coin_price and coin_price > 0:
                            price_source = f"历史价格: {coin_price:,.0f}円"
                        else:
                            price_source = "默认价格"
                        
                        formula_detail = f"""
【取引{transaction_counter}】{date} {coin}锁定奖励
奖励数量: {format_number(reward_quantity)}
奖励类型: Simple Earn Locked Rewards
价格来源: {price_source}
奖励价值（日元）: {reward_value:,.0f} 円
"""
                        formula_details.append(formula_detail)
                        
                        print(f"  → {coin}锁定奖励: 数量={reward_quantity}, 价值={reward_value:,.0f} 円")
                        
                        # 奖励作为收入记录
                        results.append([transaction_counter, date, "REWARD", coin, reward_quantity, "N/A", reward_value, "N/A"])
                        
                        # 処理済みとして記録
                        processed_transactions.add(transaction_key)

            # 結果をDataFrameに
            if results:
                self.result_df = pd.DataFrame(results, columns=["番号", "日付", "種別", "通貨", "数量", "単価(JPY)", "損益(JPY)", "平均取得単価"])

                # 年間合計損益（日元）
                total_profit = sum([float(str(r[6]).replace(',', '')) for r in results if r[6] != 0 and str(r[6]) != "N/A"])
                
                # 公式詳細を表示
                self.display_formulas(formula_details, total_profit, results)
                
                print(f"\n=== 計算結果サマリー ===")
                print(f"総取引件数: {len(results)}")
                print(f"年間損益合計: {total_profit:,.0f} 円")
                print(f"通貨別保有状況:")
                for coin, data in holdings.items():
                    if data['quantity'] > 0:
                        print(f"  {coin}: 数量={data['quantity']}, 取得費={data['total_cost']:,.0f} 円")
                
                # 結果を表示
                self.display_results()
                self.total_profit_var.set(f"年間損益合計: {total_profit:,.0f} 円")
                
                messagebox.showinfo("完了", "損益計算が完了しました！")
            else:
                messagebox.showwarning("警告", "計算可能な取引データが見つかりませんでした")
            
        except Exception as e:
            messagebox.showerror("エラー", f"計算中にエラーが発生しました:\n{str(e)}")
            print(f"エラー詳細: {e}")
    
    def calculate_jpy_only(self):
        """只计算日元交易（跳过所有ETH/BTC交易）"""
        if self.df is None:
            messagebox.showerror("エラー", "先にデータを読み込んでください")
            return
            
        try:
            # 只选择有日元价格的交易
            jpy_operations = [
                'Transaction Buy',           # BTC日元购买
                'Transaction Sold',          # BTC日元卖出
                'Transaction Spend',         # 日元支出
                'Transaction Revenue',       # 日元收入
                'Transaction Fee',           # 日元手续费
                'Simple Earn Locked Rewards' # 奖励
            ]
            
            # 只处理BTC和JPY的交易
            jpy_df = self.df[
                (self.df['Operation'].isin(jpy_operations)) & 
                (self.df['Coin'].isin(['BTC', 'JPY']))
            ].copy()
            
            if len(jpy_df) == 0:
                messagebox.showwarning("警告", "日元交易データが見つかりませんでした")
                return
            
            # 日付でソート
            jpy_df['UTC_Time'] = pd.to_datetime(jpy_df['UTC_Time'])
            jpy_df = jpy_df.sort_values('UTC_Time')
            
            print(f"日元交易データ総数: {len(jpy_df)}")
            print("交易タイプ別件数:")
            print(jpy_df['Operation'].value_counts())
            
            # 计算用变量
            holdings = {'BTC': {'quantity': 0.0, 'total_cost': 0.0}}
            results = []
            formula_details = []
            processed_transactions = set()
            transaction_counter = 0
            
            def format_number(value, decimal_places=8):
                if abs(value) < 0.000001:
                    return f"{value:.8f}"
                elif abs(value) < 0.01:
                    return f"{value:.6f}"
                elif abs(value) < 1:
                    return f"{value:.4f}"
                else:
                    return f"{value:.2f}"
            
            for idx, row in jpy_df.iterrows():
                operation = row['Operation']
                coin = row['Coin']
                change = float(row['Change'])
                date = row['UTC_Time'].strftime('%Y-%m-%d %H:%M')
                
                # 重複処理防止
                transaction_key = f"{date}_{operation}_{coin}_{change}"
                if transaction_key in processed_transactions:
                    continue
                
                print(f"処理中: {date} | {operation} | {coin} | {change}")
                
                if operation == 'Transaction Buy' and coin == 'BTC':
                    # BTC日元购买
                    if change > 0:
                        transaction_counter += 1
                        
                        # 查找相关的日元支出
                        same_time_transactions = jpy_df[jpy_df['UTC_Time'] == row['UTC_Time']]
                        jpy_spend = same_time_transactions[
                            (same_time_transactions['Operation'] == 'Transaction Spend') & 
                            (same_time_transactions['Coin'] == 'JPY')
                        ]
                        
                        jpy_spend_amount = 0.0
                        jpy_fee_amount = 0.0
                        
                        if not jpy_spend.empty:
                            jpy_spend_amount = abs(float(jpy_spend.iloc[0]['Change']))
                        
                        # 查找日元手续费
                        jpy_fees = same_time_transactions[
                            (same_time_transactions['Operation'] == 'Transaction Fee') & 
                            (same_time_transactions['Coin'] == 'JPY')
                        ]
                        
                        if not jpy_fees.empty:
                            jpy_fee_amount = abs(float(jpy_fees.iloc[0]['Change']))
                        
                        if jpy_spend_amount > 0:
                            # 更新持有量和成本
                            old_quantity = holdings['BTC']['quantity']
                            old_cost = holdings['BTC']['total_cost']
                            
                            holdings['BTC']['quantity'] += change
                            holdings['BTC']['total_cost'] += jpy_spend_amount + jpy_fee_amount
                            
                            new_quantity = holdings['BTC']['quantity']
                            new_cost = holdings['BTC']['total_cost']
                            avg_cost = new_cost / new_quantity if new_quantity > 0 else 0
                            
                            formula_detail = f"""
【取引{transaction_counter}】{date} BTC日元購入
購入数量: {format_number(change)}
購入金額（日元）: {jpy_spend_amount:,.0f} 円
手数料: {jpy_fee_amount:,.0f} 円
購入前保有数量: {format_number(old_quantity)}
購入前累計取得費: {old_cost:,.0f} 円
購入後保有数量: {format_number(new_quantity)}
購入後累計取得費: {new_cost:,.0f} 円
平均取得単価: {avg_cost:,.0f} 円
"""
                            formula_details.append(formula_detail)
                            
                            print(f"  → BTC購入: 数量={change}, 金額={jpy_spend_amount:,.0f} 円, 手数料={jpy_fee_amount:,.0f} 円")
                            
                            results.append([transaction_counter, date, "BUY", coin, change, f"{jpy_spend_amount:,.0f}", 0, f"{avg_cost:,.0f}"])
                            
                            # 标记为已处理
                            processed_transactions.add(transaction_key)
                            for _, related in jpy_spend.iterrows():
                                related_key = f"{date}_{related['Operation']}_{related['Coin']}_{related['Change']}"
                                processed_transactions.add(related_key)
                            for _, related in jpy_fees.iterrows():
                                related_key = f"{date}_{related['Operation']}_{related['Coin']}_{related['Change']}"
                                processed_transactions.add(related_key)
                
                elif operation == 'Transaction Sold' and coin == 'BTC':
                    # BTC日元卖出
                    if change < 0:
                        transaction_counter += 1
                        
                        quantity = abs(change)
                        
                        if holdings['BTC']['quantity'] >= quantity:
                            # 计算损益
                            old_quantity = holdings['BTC']['quantity']
                            old_cost = holdings['BTC']['total_cost']
                            avg_cost = old_cost / old_quantity if old_quantity > 0 else 0
                            cost_basis = avg_cost * quantity
                            
                            # 查找日元收入
                            same_time_transactions = jpy_df[jpy_df['UTC_Time'] == row['UTC_Time']]
                            jpy_revenue = same_time_transactions[
                                (same_time_transactions['Operation'] == 'Transaction Revenue') & 
                                (same_time_transactions['Coin'] == 'JPY')
                            ]
                            
                            jpy_revenue_amount = 0.0
                            jpy_fee_amount = 0.0
                            
                            if not jpy_revenue.empty:
                                jpy_revenue_amount = float(jpy_revenue.iloc[0]['Change'])
                            
                            # 查找日元手续费
                            jpy_fees = same_time_transactions[
                                (same_time_transactions['Operation'] == 'Transaction Fee') & 
                                (same_time_transactions['Coin'] == 'JPY')
                            ]
                            
                            if not jpy_fees.empty:
                                jpy_fee_amount = abs(float(jpy_fees.iloc[0]['Change']))
                            
                            if jpy_revenue_amount > 0:
                                # 计算损益
                                profit = jpy_revenue_amount - cost_basis - jpy_fee_amount
                                
                                # 更新持有量
                                new_quantity = old_quantity - quantity
                                new_cost = old_cost - cost_basis
                                
                                formula_detail = f"""
【取引{transaction_counter}】{date} BTC日元売却
売却数量: {format_number(quantity)}
売却代金（日元）: {jpy_revenue_amount:,.0f} 円
手数料: {jpy_fee_amount:,.0f} 円
売却前保有数量: {format_number(old_quantity)}
売却前累計取得費: {old_cost:,.0f} 円
平均取得単価: {avg_cost:,.0f} 円
取得費（売却分）: {cost_basis:,.0f} 円
損益計算: {jpy_revenue_amount:,.0f} 円 - {cost_basis:,.0f} 円 - {jpy_fee_amount:,.0f} 円 = {profit:,.0f} 円
売却後保有数量: {format_number(new_quantity)}
売却後累計取得費: {new_cost:,.0f} 円
"""
                                formula_details.append(formula_detail)
                                
                                print(f"  → BTC売却: 数量={quantity}, 売却代金={jpy_revenue_amount:,.0f} 円, 取得費={cost_basis:,.0f} 円, 手数料={jpy_fee_amount:,.0f} 円, 損益={profit:,.0f} 円")
                                
                                # 更新持有量
                                holdings['BTC']['quantity'] -= quantity
                                holdings['BTC']['total_cost'] -= cost_basis
                                
                                results.append([transaction_counter, date, "SELL", coin, quantity, "N/A", f"{profit:,.0f}", f"{avg_cost:,.0f}"])
                                
                                # 标记为已处理
                                processed_transactions.add(transaction_key)
                                for _, related in jpy_revenue.iterrows():
                                    related_key = f"{date}_{related['Operation']}_{related['Coin']}_{related['Change']}"
                                    processed_transactions.add(related_key)
                                for _, related in jpy_fees.iterrows():
                                    related_key = f"{date}_{related['Operation']}_{related['Coin']}_{related['Change']}"
                                    processed_transactions.add(related_key)
                        else:
                            print(f"警告: BTCの保有数量が不足しています。必要: {quantity}, 保有: {holdings['BTC']['quantity']}")
            
            # 显示结果
            if results:
                self.result_df = pd.DataFrame(results, columns=["番号", "日付", "種別", "通貨", "数量", "単価(JPY)", "損益(JPY)", "平均取得単価"])
                
                # 计算总损益
                total_profit = sum([float(str(r[6]).replace(',', '')) for r in results if r[6] != 0 and str(r[6]) != "N/A"])
                
                # 显示公式和结果
                self.display_formulas(formula_details, total_profit, results)
                
                print(f"\n=== 日元交易計算結果サマリー ===")
                print(f"総取引件数: {len(results)}")
                print(f"年間損益合計: {total_profit:,.0f} 円")
                print(f"通貨別保有状況:")
                for coin, data in holdings.items():
                    if data['quantity'] > 0:
                        print(f"  {coin}: 数量={data['quantity']}, 取得費={data['total_cost']:,.0f} 円")
                
                # 显示结果
                self.display_results()
                self.total_profit_var.set(f"年間損益合計: {total_profit:,.0f} 円")
                
                messagebox.showinfo("完了", "日元交易の損益計算が完了しました！")
            else:
                messagebox.showwarning("警告", "計算可能な日元交易データが見つかりませんでした")
            
        except Exception as e:
            messagebox.showerror("エラー", f"計算中にエラーが発生しました:\n{str(e)}")
            print(f"エラー詳細: {e}")
    
    def get_historical_prices(self):
        """自动获取ETH/BTC交易的历史日元价格"""
        if self.df is None:
            messagebox.showerror("エラー", "先にデータを読み込んでください")
            return
            
        try:
            # 查找所有ETH/BTC交易
            eth_btc_operations = [
                'Transaction Buy',           # ETH购买
                'Transaction Sold',          # ETH卖出
                'Transaction Spend',         # BTC支出
                'Transaction Revenue',       # BTC收入
            ]
            
            # 筛选ETH和BTC的交易
            eth_btc_df = self.df[
                (self.df['Operation'].isin(eth_btc_operations)) & 
                (self.df['Coin'].isin(['ETH', 'BTC']))
            ].copy()
            
            if len(eth_btc_df) == 0:
                messagebox.showwarning("警告", "ETH/BTC交易データが見つかりませんでした")
                return
            
            # 按日期排序
            eth_btc_df['UTC_Time'] = pd.to_datetime(eth_btc_df['UTC_Time'])
            eth_btc_df = eth_btc_df.sort_values('UTC_Time')
            
            print(f"ETH/BTC交易データ総数: {len(eth_btc_df)}")
            
            # 收集需要查询价格的日期
            dates_to_query = set()
            for idx, row in eth_btc_df.iterrows():
                date_str = row['UTC_Time'].strftime('%Y-%m-%d')
                dates_to_query.add(date_str)
            
            print(f"需要查询价格的日期数量: {len(dates_to_query)}")
            print("日期列表:")
            for date in sorted(dates_to_query):
                print(f"  {date}")
            
            # 显示进度对话框
            progress_window = tk.Toplevel(self.root)
            progress_window.title("正在获取历史价格")
            progress_window.geometry("400x200")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # 进度标签
            progress_label = ttk.Label(progress_window, text="正在从网络获取ETH/BTC历史价格...", font=("Arial", 12))
            progress_label.pack(pady=20)
            
            # 进度条
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
            progress_bar.pack(pady=20, padx=20, fill=tk.X)
            progress_bar.start()
            
            # 状态标签
            status_label = ttk.Label(progress_window, text="准备中...", font=("Arial", 10))
            status_label.pack(pady=10)
            
            # 在新线程中获取价格
            import threading
            def fetch_prices():
                try:
                    prices_data = {}
                    dates_list = sorted(dates_to_query)
                    
                    for i, date in enumerate(dates_list):
                        # 更新状态
                        progress_window.after(0, lambda d=date, idx=i: status_label.config(text=f"正在获取 {d} 的价格... ({idx+1}/{len(dates_list)})"))
                        
                        # 获取ETH价格
                        eth_price = self.fetch_eth_price_for_date(date)
                        # 获取BTC价格
                        btc_price = self.fetch_btc_price_for_date(date)
                        
                        if eth_price > 0 or btc_price > 0:
                            prices_data[date] = {
                                'ETH': eth_price,
                                'BTC': btc_price
                            }
                        
                        # 短暂延迟避免请求过快
                        import time
                        time.sleep(0.5)
                    
                    # 保存价格数据
                    if prices_data:
                        json_filename = "historical_prices.json"
                        with open(json_filename, 'w', encoding='utf-8') as f:
                            import json
                            json.dump(prices_data, f, ensure_ascii=False, indent=2)
                        
                        # 显示成功消息
                        progress_window.after(0, lambda: self.show_price_result(prices_data, json_filename, progress_window))
                    else:
                        progress_window.after(0, lambda: self.show_price_error("未能获取到任何价格数据", progress_window))
                        
                except Exception as e:
                    progress_window.after(0, lambda: self.show_price_error(f"获取价格时发生错误: {str(e)}", progress_window))
            
            # 启动线程
            thread = threading.Thread(target=fetch_prices)
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            messagebox.showerror("エラー", f"历史价格获取中にエラーが発生しました:\n{str(e)}")
            print(f"エラー詳細: {e}")
    
    def fetch_eth_price_for_date(self, date_str):
        """从网络获取指定日期的ETH价格（日元）"""
        try:
            import requests
            from datetime import datetime
            
            # 解析日期
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            timestamp = int(date_obj.timestamp())
            
            # 使用CoinGecko API获取ETH价格
            url = f"https://api.coingecko.com/api/v3/coins/ethereum/market_chart/range"
            params = {
                'vs_currency': 'jpy',
                'from': timestamp,
                'to': timestamp + 86400  # 加24小时
            }
            
            response = requests.get(url, params=params, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if 'prices' in data and len(data['prices']) > 0:
                    # 取第一个价格（最接近指定日期的）
                    price = data['prices'][0][1]
                    print(f"  ✓ {date_str} ETH价格: {price:,.0f}円")
                    return price
            
            # 如果CoinGecko失败，尝试使用CoinMarketCap（需要API key）
            # 这里暂时返回0，你可以添加自己的API key
            
            print(f"  ✗ {date_str} ETH价格获取失败")
            return 0
            
        except Exception as e:
            print(f"  ✗ {date_str} ETH价格获取错误: {e}")
            return 0
    
    def fetch_btc_price_for_date(self, date_str):
        """从网络获取指定日期的BTC价格（日元）"""
        try:
            import requests
            from datetime import datetime
            
            # 解析日期
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            timestamp = int(date_obj.timestamp())
            
            # 使用CoinGecko API获取BTC价格
            url = f"https://api.coingecko.com/api/v3/coins/bitcoin/market_chart/range"
            params = {
                'vs_currency': 'jpy',
                'from': timestamp,
                'to': timestamp + 86400  # 加24小时
            }
            
            response = requests.get(url, params=params, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if 'prices' in data and len(data['prices']) > 0:
                    # 取第一个价格（最接近指定日期的）
                    price = data['prices'][0][1]
                    print(f"  ✓ {date_str} BTC价格: {price:,.0f}円")
                    return price
            
            print(f"  ✗ {date_str} BTC价格获取失败")
            return 0
            
        except Exception as e:
            print(f"  ✗ {date_str} BTC价格获取错误: {e}")
            return 0
    
    def show_price_result(self, prices_data, filename, window):
        """显示价格获取结果"""
        window.destroy()
        
        messagebox.showinfo("成功", f"历史价格获取完成！\n\n共获取了 {len(prices_data)} 个日期的价格\n已保存到: {filename}\n\n现在可以使用'损益计算执行'按钮进行更准确的计算！")
        
        print(f"\n=== 历史价格获取完成 ===")
        print(f"保存文件: {filename}")
        print("获取的价格数据:")
        for date, prices in prices_data.items():
            print(f"  {date}: ETH={prices['ETH']:,.0f}円, BTC={prices['BTC']:,.0f}円")
    
    def show_price_error(self, error_msg, window):
        """显示价格获取错误"""
        window.destroy()
        messagebox.showerror("错误", error_msg)
    

    
    def load_historical_prices(self):
        """从JSON文件加载历史价格"""
        try:
            json_filename = "historical_prices.json"
            if os.path.exists(json_filename):
                with open(json_filename, 'r', encoding='utf-8') as f:
                    import json
                    return json.load(f)
            return {}
        except Exception as e:
            print(f"历史价格加载エラー: {e}")
            return {}
    
    def get_price_for_date(self, date_str, coin):
        """根据日期和币种获取价格"""
        prices = self.load_historical_prices()
        date_key = date_str[:10]  # 只取日期部分，去掉时间
        
        if date_key in prices and coin in prices[date_key]:
            return prices[date_key][coin]
        return None
    
    def display_formulas(self, formula_details, total_profit, results):
        """公式詳細を表示"""
        self.formula_text.config(state=tk.NORMAL)
        self.formula_text.delete(1.0, tk.END)
        
        # 総合計公式
        total_formula = f"""
【年間損益合計の計算公式】
"""
        
        # 各取引の損益を抽出
        sell_results = [r for r in results if r[2] == "SELL"]  # 番号列が追加されたため、インデックスを調整
        if sell_results:
            total_formula += "年間損益合計 = "
            profit_terms = []
            for result in sell_results:
                transaction_num = result[0]  # 取引番号
                profit = result[6]  # 損益(BTC)のインデックスを調整
                profit_terms.append(f"取引{transaction_num}の損益({profit})")
            
            total_formula += " + ".join(profit_terms)
            total_profit = sum([float(str(r[6]).replace(',', '')) for r in sell_results if r[6] != 0 and str(r[6]) != "N/A"])
            total_formula += f" = {total_profit:,.0f} 円"
        
        self.formula_text.insert(tk.END, total_formula)
        
        # 各取引の詳細公式
        for i, detail in enumerate(formula_details):
            self.formula_text.insert(tk.END, f"\n{'='*50}\n{detail}")
        
        self.formula_text.config(state=tk.DISABLED)
            
    def display_results(self):
        # 既存の結果をクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # 新しい結果を表示
        for _, row in self.result_df.iterrows():
            self.tree.insert("", tk.END, values=list(row))

def main():
    root = tk.Tk()
    app = CryptoCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
