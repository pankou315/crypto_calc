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
        self.trading_df = None  # 現物注文取引履歴用
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
        
        # 第一个文件：取りれ歴トランザクション記録
        ttk.Label(file_frame, text="取りれ歴トランザクション記録:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=80)
        file_entry.grid(row=1, column=0, padx=(0, 10), pady=(0, 10))
        
        browse_btn = ttk.Button(file_frame, text="ファイル選択", command=self.browse_file)
        browse_btn.grid(row=1, column=1, pady=(0, 10))
        
        load_btn = ttk.Button(file_frame, text="データ読み込み", command=self.load_data)
        load_btn.grid(row=1, column=2, padx=(10, 0), pady=(0, 10))
        
        # 第二个文件：現物注文取引履歴
        ttk.Label(file_frame, text="現物注文取引履歴:").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        self.trading_file_path_var = tk.StringVar()
        trading_file_entry = ttk.Entry(file_frame, textvariable=self.trading_file_path_var, width=80)
        trading_file_entry.grid(row=3, column=0, padx=(0, 10), pady=(0, 10))
        
        trading_browse_btn = ttk.Button(file_frame, text="ファイル選択", command=self.browse_trading_file)
        trading_browse_btn.grid(row=3, column=1, pady=(0, 10))
        
        trading_load_btn = ttk.Button(file_frame, text="データ読み込み", command=self.load_trading_data)
        trading_load_btn.grid(row=3, column=2, padx=(10, 0), pady=(0, 10))
        
        # アルゴリズム説明部分
        algo_frame = ttk.LabelFrame(main_frame, text="計算アルゴリズム", padding="10")
        algo_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # 计算方法选择
        method_frame = ttk.Frame(algo_frame)
        method_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(method_frame, text="計算方法選択:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.calculation_method = tk.StringVar(value="moving_average")
        moving_avg_radio = ttk.Radiobutton(method_frame, text="移動平均法", variable=self.calculation_method, value="moving_average")
        moving_avg_radio.grid(row=0, column=1, padx=(0, 20))
        
        total_avg_radio = ttk.Radiobutton(method_frame, text="総平均法", variable=self.calculation_method, value="total_average")
        total_avg_radio.grid(row=0, column=2, padx=(0, 20))
        
        # 方法说明
        method_desc = ttk.Label(method_frame, text="移動平均法: 税務署推奨 | 総平均法: 簡易計算", font=("Arial", 9))
        method_desc.grid(row=0, column=3, padx=(20, 0))
        
        algo_text = tk.Text(algo_frame, height=6, width=120, wrap=tk.WORD)
        algo_text.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
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

【総平均法による損益計算アルゴリズム】

1. 年間の総取得費と総取得数量を計算
2. 総平均取得単価 = 総取得費 ÷ 総取得数量
3. 各売却取引の損益 = 売却代金 - (売却数量 × 総平均取得単価) - 手数料
4. 年間損益合計 = 全売却取引の損益合計

※ 移動平均法は税務署が推奨する方法です
※ 総平均法は簡易計算方法です

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

【現物注文取引履歴連携機能】
• 現物注文取引履歴.csvファイルを読み込むことで、実際の取引価格を使用
• 取引価格の優先順位：現物注文取引履歴 > 历史价格 > デフォルト価格
• より正確な損益計算が可能

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
        
        # 重新生成文件按钮
        regenerate_btn = ttk.Button(main_frame, text="重新生成文件", command=self.regenerate_files)
        regenerate_btn.grid(row=3, column=3, pady=20)
        
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
            
    def browse_trading_file(self):
        filename = filedialog.askopenfilename(
            title="現物注文取引履歴CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.trading_file_path_var.set(filename)
            
    def load_trading_data(self):
        file_path = self.trading_file_path_var.get()
        if not file_path:
            messagebox.showerror("エラー", "現物注文取引履歴ファイルを選択してください")
            return
            
        try:
            # 文字コードを自動判定
            encodings = ['utf-8', 'shift-jis', 'cp932']
            for encoding in encodings:
                try:
                    self.trading_df = pd.read_csv(file_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                messagebox.showerror("エラー", "ファイルの文字コードを判定できませんでした")
                return
                
            # データの基本情報を表示
            messagebox.showinfo("成功", f"現物注文取引履歴データを読み込みました\n行数: {len(self.trading_df)}\n列: {', '.join(self.trading_df.columns)}")
            
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました: {str(e)}")
            
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
                
            # 自動的に現物注文取引履歴.csvファイルを読み込む
            trading_file_path = "現物注文取引履歴.csv"
            if os.path.exists(trading_file_path):
                try:
                    # 文字コードを自動判定
                    for encoding in encodings:
                        try:
                            self.trading_df = pd.read_csv(trading_file_path, encoding=encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        print("警告: 現物注文取引履歴.csvファイルの文字コードを判定できませんでした")
                        self.trading_df = None
                        
                    if self.trading_df is not None:
                        print(f"✓ 自動的に現物注文取引履歴.csvファイルを読み込みました (行数: {len(self.trading_df)})")
                except Exception as e:
                    print(f"警告: 現物注文取引履歴.csvファイルの読み込みに失敗しました: {str(e)}")
                    self.trading_df = None
            else:
                print("警告: 現物注文取引履歴.csvファイルが見つかりません")
                self.trading_df = None
            
            messagebox.showinfo("成功", f"データを読み込みました\n行数: {len(self.df)}\n列数: {len(self.df.columns)}")
            
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました:\n{str(e)}")
            
    def process_trading_history(self):
        """現物注文取引履歴データを処理して、取引価格情報を追加"""
        try:
            print("現物注文取引履歴データの処理を開始...")
            
            # 检查trading_df是否存在
            if self.trading_df is None:
                print("警告: 現物注文取引履歴データが読み込まれていません")
                return
            
            # 日付列をdatetime型に変換
            self.trading_df['Date(UTC)'] = pd.to_datetime(self.trading_df['Date(UTC)'])
            
            # 取引データを日付でソート
            self.trading_df = self.trading_df.sort_values('Date(UTC)')
            
            # 各取引ペアの価格情報を整理
            self.trading_prices = {}
            
            processed_count = 0
            error_count = 0
            
            for idx, row in self.trading_df.iterrows():
                try:
                    date = row['Date(UTC)']
                    pair = row['Pair']
                    side = row['Side']
                    price = float(row['Price'])
                    executed = row['Executed']
                    amount = row['Amount']
                    fee = row['Fee']
                    
                    # 取引ペアを解析（例：BTCJPY, ETHBTC, DOGEJPY）
                    if 'JPY' in pair:
                        # 日元取引の場合
                        base_coin = pair.replace('JPY', '')
                        if side == 'BUY':
                            # 購入：JPY支出、暗号資産獲得
                            if base_coin not in self.trading_prices:
                                self.trading_prices[base_coin] = {'buys': [], 'sells': []}
                            
                            # 数量と金額を解析
                            coin_quantity = self.parse_quantity(executed)
                            jpy_amount = self.parse_jpy_amount(amount)
                            
                            if coin_quantity > 0:  # 有効な数量の場合のみ追加
                                self.trading_prices[base_coin]['buys'].append({
                                    'date': date,
                                    'price': price,
                                    'quantity': coin_quantity,
                                    'jpy_amount': jpy_amount,
                                    'fee': fee,
                                    'pair': pair  # 添加交易对信息
                                })
                                processed_count += 1
                            
                        elif side == 'SELL':
                            # 売却：暗号資産支出、JPY獲得
                            if base_coin not in self.trading_prices:
                                self.trading_prices[base_coin] = {'buys': [], 'sells': []}
                            
                            coin_quantity = self.parse_quantity(executed)
                            jpy_amount = self.parse_jpy_amount(amount)
                            
                            if coin_quantity > 0:  # 有効な数量の場合のみ追加
                                self.trading_prices[base_coin]['sells'].append({
                                    'date': date,
                                    'price': price,
                                    'quantity': coin_quantity,
                                    'jpy_amount': jpy_amount,
                                    'fee': fee,
                                    'pair': pair  # 添加交易对信息
                                })
                                processed_count += 1
                            
                    elif 'BTC' in pair and 'ETH' in pair:
                        # ETHBTC取引の場合 - 只记录BTC的价格
                        if 'BTC' not in self.trading_prices:
                            self.trading_prices['BTC'] = {'buys': [], 'sells': []}
                        
                        eth_quantity = self.parse_quantity(executed)
                        btc_amount = self.parse_btc_amount(amount)
                        
                        if btc_amount > 0:  # 有効なBTC数量の場合のみ追加
                            if side == 'BUY':
                                # ETH購入（BTC支出）→ BTC売却
                                self.trading_prices['BTC']['sells'].append({
                                    'date': date,
                                    'price': 1/price,  # BTC/ETH価格
                                    'quantity': btc_amount,
                                    'eth_amount': eth_quantity,
                                    'fee': fee,
                                    'pair': pair
                                })
                                processed_count += 1
                            elif side == 'SELL':
                                # ETH売却（BTC獲得）→ BTC購入
                                self.trading_prices['BTC']['buys'].append({
                                    'date': date,
                                    'price': 1/price,  # BTC/ETH価格
                                    'quantity': btc_amount,
                                    'eth_amount': eth_quantity,
                                    'fee': fee,
                                    'pair': pair
                                })
                                processed_count += 1
                                
                    elif 'ETH' in pair and 'DAI' in pair:
                        # ETHDAI取引の場合
                        if 'ETH' not in self.trading_prices:
                            self.trading_prices['ETH'] = {'buys': [], 'sells': []}
                        
                        eth_quantity = self.parse_quantity(executed)
                        dai_amount = self.parse_dai_amount(amount)
                        
                        if eth_quantity > 0:  # 有効な数量の場合のみ追加
                            if side == 'BUY':
                                # ETH購入（DAI支出）
                                self.trading_prices['ETH']['buys'].append({
                                    'date': date,
                                    'price': price,  # ETH/DAI価格
                                    'quantity': eth_quantity,
                                    'dai_amount': dai_amount,
                                    'fee': fee
                                })
                                processed_count += 1
                            elif side == 'SELL':
                                # ETH売却（DAI獲得）
                                self.trading_prices['ETH']['sells'].append({
                                    'date': date,
                                    'price': price,
                                    'quantity': eth_quantity,
                                    'dai_amount': dai_amount,
                                    'fee': fee
                                })
                                processed_count += 1
                                
                except Exception as e:
                    error_count += 1
                    print(f"行 {idx} の処理エラー: {str(e)}")
                    print(f"  データ: {row.to_dict()}")
                    continue
            
            print(f"現物注文取引履歴処理完了:")
            print(f"  ✓ 処理成功: {processed_count} 件")
            print(f"  ⚠️ 処理エラー: {error_count} 件")
            print(f"  ✓ 通貨数: {len(self.trading_prices)}")
            
            # 各通貨の取引件数を表示
            for coin, data in self.trading_prices.items():
                buy_count = len(data.get('buys', []))
                sell_count = len(data.get('sells', []))
                print(f"    {coin}: 購入 {buy_count}件, 売却 {sell_count}件")
            
        except Exception as e:
            print(f"現物注文取引履歴処理エラー: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("エラー", f"現物注文取引履歴の処理に失敗しました: {str(e)}")
    
    def parse_quantity(self, executed_str):
        """実行数量文字列から数量を抽出"""
        try:
            # 例: "0.003ETH" -> 0.003
            for char in executed_str:
                if char.isalpha():
                    return float(executed_str.replace(char, ''))
            return 0.0
        except:
            return 0.0
            
    def parse_jpy_amount(self, amount_str):
        """金額文字列からJPY金額を抽出"""
        try:
            if 'JPY' in amount_str:
                return float(amount_str.replace('JPY', ''))
            return 0.0
        except:
            return 0.0
            
    def parse_btc_amount(self, amount_str):
        """金額文字列からBTC金額を抽出"""
        try:
            if 'BTC' in amount_str:
                return float(amount_str.replace('BTC', ''))
            return 0.0
        except:
            return 0.0
            
    def parse_dai_amount(self, amount_str):
        """金額文字列からDAI金額を抽出"""
        try:
            if 'DAI' in amount_str:
                return float(amount_str.replace('DAI', ''))
            return 0.0
        except:
            return 0.0
            
    def get_trading_price(self, coin, date, operation):
        """指定した日付と通貨の取引価格を取得"""
        if coin not in self.trading_prices:
            return None
            
        # 日付をdatetime型に変換
        if isinstance(date, str):
            date = pd.to_datetime(date)
            
        # 最も近い日付の取引価格を検索
        if operation == 'BUY':
            trades = self.trading_prices[coin]['buys']
        else:
            trades = self.trading_prices[coin]['sells']
            
        if not trades:
            return None
            
        # 日付の差が最小の取引を検索
        min_diff = float('inf')
        best_price = None
        
        for trade in trades:
            diff = abs((trade['date'] - date).total_seconds())
            if diff < min_diff:
                min_diff = diff
                best_price = trade['price']
                
        return best_price
            
    def calculate_profit(self):
        if self.df is None:
            messagebox.showerror("エラー", "先にデータを読み込んでください")
            return
            
        if self.trading_df is None:
            messagebox.showerror("エラー", "先に現物注文取引履歴データを読み込んでください")
            return
            
        try:
            # 获取选择的计算方法
            method = self.calculation_method.get()
            print(f"=== 使用计算方法: {'移動平均法' if method == 'moving_average' else '総平均法'} ===")
            
            # 現物注文取引履歴データを処理
            self.process_trading_history()
            
            print(f"=== 現物注文取引履歴データ処理完了 ===")
            print(f"利用可能な通貨: {list(self.trading_prices.keys())}")
            for coin, data in self.trading_prices.items():
                print(f"  {coin}: 購入{len(data['buys'])}件, 売却{len(data['sells'])}件")
            
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
            
            # 总平均法计算用变量
            total_buy_quantity = 0.0
            total_buy_cost = 0.0
            total_avg_cost = 0.0
            
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

            # 現物注文取引履歴から価格を取得する関数
            def get_price_from_trading_history(coin, date, operation_type):
                """現物注文取引履歴から指定した日付と通貨の価格を取得"""
                if not hasattr(self, 'trading_prices') or coin not in self.trading_prices:
                    return None
                
                try:
                    # 日付をdatetime型に変換
                    if isinstance(date, str):
                        date_obj = pd.to_datetime(date)
                    else:
                        date_obj = date
                    
                    # 最も近い日付の取引価格を検索
                    min_diff = float('inf')
                    best_price = None
                    best_trade_info = None
                    
                    # 取引タイプに応じて検索
                    if operation_type == 'BUY':
                        trades = self.trading_prices[coin]['buys']
                    else:  # SELL
                        trades = self.trading_prices[coin]['sells']
                    
                    if not trades:
                        return None
                    
                    for trade in trades:
                        diff = abs((trade['date'] - date_obj).total_seconds())
                        if diff < min_diff:
                            min_diff = diff
                            best_price = trade['price']
                            best_trade_info = f"{operation_type} {trade['date']} {trade['price']}"
                    
                    # 2時間以内の価格のみ有効とする
                    if best_price is not None and min_diff <= 7200:
                        print(f"    ✓ 現物注文取引履歴から{coin}の{operation_type}価格を取得: {best_price} (日付差: {min_diff/60:.1f}分)")
                        return best_price
                    elif best_price is not None:
                        print(f"    ⚠️ 現物注文取引履歴から{coin}の{operation_type}価格を取得: {best_price} (日付差: {min_diff/60:.1f}分) - 日付差が大きすぎます")
                    
                    return None
                    
                except Exception as e:
                    print(f"    現物注文取引履歴からの価格取得エラー: {str(e)}")
                    return None
                    
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
                            # 更新持有量和成本
                            old_quantity = holdings['BTC']['quantity']
                            old_cost = holdings['BTC']['total_cost']
                            
                            holdings['BTC']['quantity'] += change
                            holdings['BTC']['total_cost'] += jpy_spend_amount + jpy_fee_amount
                            
                            new_quantity = holdings['BTC']['quantity']
                            new_cost = holdings['BTC']['total_cost']
                            avg_cost = new_cost / new_quantity if new_quantity > 0 else 0
                            
                            # 更新总平均法变量
                            total_buy_quantity += change
                            total_buy_cost += jpy_spend_amount + jpy_fee_amount
                            total_avg_cost = total_buy_cost / total_buy_quantity if total_buy_quantity > 0 else 0
                            
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
                        # 重複処理防止 - 使用更精确的key来避免重复计算
                        transaction_key = f"{date}_{operation}_{coin}_{change}"
                        if transaction_key in processed_transactions:
                            print(f"重複{coin} SELL取引をスキップ: {transaction_key}")
                            continue
                        
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
                                # 首先尝试从現物注文取引履歴获取实际卖出价格
                                actual_sell_price = None
                                if hasattr(self, 'trading_df') and self.trading_df is not None:
                                    try:
                                        # 查找对应日期的交易价格
                                        transaction_date = row['UTC_Time'].strftime('%Y-%m-%d %H:%M')
                                        actual_sell_price = self.get_price_for_date(transaction_date, coin)
                                        if actual_sell_price and actual_sell_price > 0:
                                            print(f"  ✓ 現物注文取引履歴から{coin}売却単価を取得: {actual_sell_price:,.0f} 円")
                                            # 使用实际价格重新计算卖出金额
                                            jpy_revenue_amount = quantity * actual_sell_price
                                            print(f"  ✓ 使用实际价格重新计算: {quantity:.8f} × {actual_sell_price:,.0f} = {jpy_revenue_amount:,.0f} 円")
                                        else:
                                            print(f"  ⚠️ 現物注文取引履歴から{coin}の価格を取得できませんでした")
                                    except Exception as e:
                                        print(f"  現物注文取引履歴からの価格取得エラー: {str(e)}")
                                
                                # 損益計算（手数料を考慮）
                                if method == "total_average":
                                    # 总平均法：使用总平均成本计算
                                    cost_basis = total_avg_cost * quantity
                                    print(f"  総平均法: 売却数量={quantity:.8f}, 総平均単価={total_avg_cost:,.0f} 円, 取得費={cost_basis:,.0f} 円")
                                else:
                                    # 移动平均法：使用当前平均成本计算
                                    cost_basis = avg_cost * quantity
                                    print(f"  移動平均法: 売却数量={quantity:.8f}, 現在平均単価={avg_cost:,.0f} 円, 取得費={cost_basis:,.0f} 円")
                                
                                profit = jpy_revenue_amount - cost_basis - jpy_fee_amount
                                
                                # 公式詳細を記録
                                new_quantity = old_quantity - quantity
                                new_cost = old_cost - cost_basis
                                
                                if method == "total_average":
                                    method_name = "総平均法"
                                    cost_calculation = f"取得費（売却分）: {total_avg_cost:,.0f} 円 × {format_number(quantity)} = {cost_basis:,.0f} 円"
                                else:
                                    method_name = "移動平均法"
                                    cost_calculation = f"取得費（売却分）: {avg_cost:,.0f} 円 × {format_number(quantity)} = {cost_basis:,.0f} 円"
                                
                                formula_detail = f"""
【取引{transaction_counter}】{date} BTC日元売却 ({method_name})
売却数量: {format_number(quantity)}
売却代金（日元）: {jpy_revenue_amount:,.0f} 円
手数料: {jpy_fee_amount:,.0f} 円
売却前保有数量: {format_number(old_quantity)}
売却前累計取得費: {old_cost:,.0f} 円
{cost_calculation}
損益計算: {jpy_revenue_amount:,.0f} 円 - {cost_basis:,.0f} 円 - {jpy_fee_amount:,.0f} 円 = {profit:,.0f} 円
売却後保有数量: {format_number(new_quantity)}
売却後累計取得費: {new_cost:,.0f} 円
"""
                                formula_details.append(formula_detail)
                                
                                print(f"  → {coin}売却: 数量={quantity}, 売却代金={jpy_revenue_amount:,.0f} 円, 取得費={cost_basis:,.0f} 円, 手数料={jpy_fee_amount:,.0f} 円, 損益={profit:,.0f} 円")
                                
                                # 更详细的卖出信息显示
                                if method == "total_average":
                                    print(f"    【総平均法計算詳細】")
                                    print(f"    売却数量: {quantity:.8f} BTC")
                                    print(f"    売却代金: {jpy_revenue_amount:,.0f} 円")
                                    print(f"    総平均単価: {total_avg_cost:,.0f} 円/BTC")
                                    print(f"    売却成本: {cost_basis:,.0f} 円 (={quantity:.8f} × {total_avg_cost:,.0f})")
                                    print(f"    手数料: {jpy_fee_amount:,.0f} 円")
                                    print(f"    損益計算: {jpy_revenue_amount:,.0f} - {cost_basis:,.0f} - {jpy_fee_amount:,.0f} = {profit:,.0f} 円")
                                else:
                                    print(f"    【移動平均法計算詳細】")
                                    print(f"    売却数量: {quantity:.8f} BTC")
                                    print(f"    売却代金: {jpy_revenue_amount:,.0f} 円")
                                    print(f"    現在平均単価: {avg_cost:,.0f} 円/BTC")
                                    print(f"    売却成本: {cost_basis:,.0f} 円 (={quantity:.8f} × {avg_cost:,.0f})")
                                    print(f"    手数料: {jpy_fee_amount:,.0f} 円")
                                    print(f"    損益計算: {jpy_revenue_amount:,.0f} - {cost_basis:,.0f} - {jpy_fee_amount:,.0f} = {profit:,.0f} 円")
                                
                                # 保有数量と取得費を減らす
                                holdings[coin]['quantity'] -= quantity
                                holdings[coin]['total_cost'] -= cost_basis
                                
                                # 売却単価を計算
                                sell_unit_price = "N/A"
                                if actual_sell_price and actual_sell_price > 0:
                                    sell_unit_price = f"{actual_sell_price:,.0f}"
                                elif quantity > 0:
                                    # 如果没有实际价格，使用计算出的单价
                                    calculated_price = jpy_revenue_amount / quantity
                                    sell_unit_price = f"{calculated_price:,.0f}"
                                
                                results.append([transaction_counter, date, "SELL", coin, quantity, sell_unit_price, f"{profit:,.0f}", f"{avg_cost:,.0f}"])
                                
                                # 処理済みとして記録 - 立即标记所有相关交易为已处理
                                processed_transactions.add(transaction_key)
                                for related in related_transactions:
                                    related_key = f"{date}_{related['Operation']}_{related['Coin']}_{related['Change']}"
                                    processed_transactions.add(related_key)
                                    print(f"    関連取引を処理済みとして記録: {related_key}")
                                
                                print(f"    ✓ SELL取引完了: {transaction_key}")
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
                            # 如果没有历史价格，使用默认价格或设为0
                            if coin == 'ETH':
                                reward_value = reward_quantity * 17.5  # 1 ETH = 17.5日元
                                print(f"  → {coin}锁定奖励: 数量={reward_quantity}, 默认价格, 价值={reward_value:,.0f}円")
                            elif coin == 'BTC':
                                reward_value = reward_quantity * 50  # 1 BTC = 50日元
                                print(f"  → {coin}锁定奖励: 数量={reward_quantity}, 默认价格, 价值={reward_value:,.0f}円")
                            else:
                                # 其他货币如果没有历史价格，设为0
                                reward_value = 0
                                print(f"  → {coin}锁定奖励: 数量={reward_quantity}, 无历史价格, 价值=0円")
                        
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
                
                # 计算买卖汇总信息
                total_buy_quantity = 0.0
                total_buy_amount = 0.0
                total_sell_quantity = 0.0
                total_sell_amount = 0.0
                
                print(f"\n=== 详细金额计算 ===")
                
                for result in results:
                    if result[2] == "BUY":  # 种别列
                        quantity = float(str(result[4]).replace(',', ''))  # 数量列
                        unit_price = str(result[5])  # 単価(JPY)列
                        
                        total_buy_quantity += quantity
                        
                        if unit_price != "N/A":
                            try:
                                amount = float(unit_price.replace(',', ''))
                                # 修复：计算总金额，而不是只加单价
                                total_amount = quantity * amount
                                total_buy_amount += total_amount
                                print(f"  BUY: {quantity:.8f} BTC × {amount:,.0f} 円 = {total_amount:,.0f} 円")
                            except:
                                print(f"  BUY: {quantity:.8f} BTC × {unit_price} (无法解析)")
                        else:
                            print(f"  BUY: {quantity:.8f} BTC × N/A (单价缺失)")
                            
                    elif result[2] == "SELL":  # 种别列
                        quantity = float(str(result[4]).replace(',', ''))  # 数量列
                        total_sell_quantity += quantity
                        
                        # 对于SELL，我们需要从損益(JPY)列计算金额
                        if result[6] != "N/A":
                            try:
                                profit = float(str(result[6]).replace(',', ''))
                                avg_cost = float(str(result[7]).replace(',', ''))
                                
                                # 修复：使用真实的卖出价格计算卖出金额
                                # 从単価(JPY)列获取真实价格
                                unit_price_str = str(result[5])
                                if unit_price_str != "N/A":
                                    try:
                                        real_unit_price = float(unit_price_str.replace(',', ''))
                                        
                                        # 添加价格合理性检查
                                        if real_unit_price > 50000000:  # 如果单价超过5000万日元，认为异常
                                            print(f"  ⚠️ 检测到异常高单价: {real_unit_price:,.0f} 円，使用回退方法")
                                            sell_amount = profit + (quantity * avg_cost)
                                            print(f"  SELL: {quantity:.8f} BTC, 損益: {profit:,.0f} 円, 取得費: {quantity * avg_cost:,.0f} 円, 売却金額: {sell_amount:,.0f} 円 (回退方法)")
                                        else:
                                            sell_amount = quantity * real_unit_price
                                            print(f"  SELL: {quantity:.8f} BTC, 損益: {profit:,.0f} 円, 取得費: {quantity * avg_cost:,.0f} 円, 売却金額: {sell_amount:,.0f} 円 (単価: {real_unit_price:,.0f} 円)")
                                    except:
                                        # 如果无法解析单价，回退到原来的方法
                                        sell_amount = profit + (quantity * avg_cost)
                                        print(f"  SELL: {quantity:.8f} BTC, 損益: {profit:,.0f} 円, 取得費: {quantity * avg_cost:,.0f} 円, 売却金額: {sell_amount:,.0f} 円 (回退方法)")
                                else:
                                    # 如果单价是N/A，回退到原来的方法
                                    sell_amount = profit + (quantity * avg_cost)
                                    print(f"  SELL: {quantity:.8f} BTC, 損益: {profit:,.0f} 円, 取得費: {quantity * avg_cost:,.0f} 円, 売却金額: {sell_amount:,.0f} 円 (回退方法)")
                                
                                total_sell_amount += sell_amount
                            except:
                                print(f"  SELL: {quantity:.8f} BTC (无法计算売却金額)")
                        else:
                            print(f"  SELL: {quantity:.8f} BTC (損益N/A)")
                
                print(f"\n=== 汇总计算结果 ===")
                print(f"总买入数量: {total_buy_quantity:.8f} BTC")
                print(f"总买入金额: {total_buy_amount:,.0f} 円")
                print(f"总卖出数量: {total_sell_quantity:.8f} BTC")
                print(f"总卖出金额: {total_sell_amount:,.0f} 円")
                if total_buy_quantity > 0:
                    print(f"平均买入单价: {total_buy_amount / total_buy_quantity:,.0f} 円/BTC")
                
                # 公式詳細を表示
                self.display_formulas(formula_details, total_profit, results)
                
                print(f"\n=== 計算結果サマリー ===")
                print(f"総取引件数: {len(results)}")
                print(f"総買入 BTC: {total_buy_quantity:.8f}, 総価格: {total_buy_amount:,.0f} 円")
                print(f"総売出 BTC: {total_sell_quantity:.8f}, 総価格: {total_sell_amount:,.0f} 円")
                print(f"年間損益合計: {total_profit:,.0f} 円")
                print(f"通貨別保有状況:")
                for coin, data in holdings.items():
                    if data['quantity'] > 0:
                        print(f"  {coin}: 数量={data['quantity']}, 取得費={data['total_cost']:,.0f} 円")
                
                # 結果を表示
                self.display_results()
                
                # 更新汇总信息显示 - 使用更详细的格式
                avg_buy_price = total_buy_amount / total_buy_quantity if total_buy_quantity > 0 else 0
                avg_sell_price = total_sell_amount / total_sell_quantity if total_sell_quantity > 0 else 0
                
                # 获取计算方法
                method = self.calculation_method.get()
                method_name = "総平均法" if method == "total_average" else "移動平均法"
                
                summary_text = f"計算方法: {method_name}\n"
                summary_text += f"総買入 BTC: {total_buy_quantity:.8f}, 単価: {avg_buy_price:,.0f} 円, 花费: {total_buy_amount:,.0f} 円\n"
                summary_text += f"総売出 BTC: {total_sell_quantity:.8f}, 単価: {avg_sell_price:,.0f} 円, 获得: {total_sell_amount:,.0f} 円\n"
                
                # 添加总平均法的详细计算说明
                if method == "total_average":
                    total_avg_cost = total_buy_amount / total_buy_quantity if total_buy_quantity > 0 else 0
                    total_sell_cost = total_sell_quantity * total_avg_cost
                    total_fees = total_sell_amount - total_sell_cost - total_profit
                    summary_text += f"\n【総平均法計算詳細】\n"
                    summary_text += f"総平均単価: {total_avg_cost:,.0f} 円/BTC\n"
                    summary_text += f"総売出成本: {total_sell_cost:,.0f} 円 (={total_sell_quantity:.8f} × {total_avg_cost:,.0f})\n"
                    summary_text += f"総手数料: {total_fees:,.0f} 円\n"
                    summary_text += f"損益計算: {total_sell_amount:,.0f} - {total_sell_cost:,.0f} - {total_fees:,.0f} = {total_profit:,.0f} 円\n"
                
                summary_text += f"\n年間損益合計: {total_profit:,.0f} 円"
                
                self.total_profit_var.set(summary_text)
                
                # 按货币分别保存结果
                results_folder = self.save_results_by_coin(results)
                
                # 生成总的税务申报表格
                if results_folder:
                    self._results_folder = results_folder
                self.generate_tax_report(results)
                
                messagebox.showinfo("完了", "損益計算が完了しました！\n\n各通貨の結果を個別に保存し、\n税務申告用の総合計算表も生成しました。")
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
            # 获取选择的计算方法
            method = self.calculation_method.get()
            print(f"=== 使用计算方法: {'移動平均法' if method == 'moving_average' else '総平均法'} ===")
            
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
            
            # 总平均法计算用变量
            total_buy_quantity = 0.0
            total_buy_cost = 0.0
            total_avg_cost = 0.0
            
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
                            
                            # 更新总平均法变量
                            total_buy_quantity += change
                            total_buy_cost += jpy_spend_amount + jpy_fee_amount
                            total_avg_cost = total_buy_cost / total_buy_quantity if total_buy_quantity > 0 else 0
                            
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
                        # 重複処理防止 - 对于BTC，使用时间作为key，避免同一时间的多个SELL交易重复计算
                        if coin == 'BTC':
                            # BTC交易：使用时间作为key，避免同一时间的多个SELL交易重复计算
                            transaction_key = f"{date}_{operation}_{coin}"
                            if transaction_key in processed_transactions:
                                print(f"重複BTC SELL取引をスキップ: {transaction_key}")
                                continue
                        else:
                            # 其他货币：使用更精确的key
                            transaction_key = f"{date}_{operation}_{coin}_{change}"
                            if transaction_key in processed_transactions:
                                print(f"重複SELL取引をスキップ: {transaction_key}")
                                continue
                        
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
                                # 首先尝试从現物注文取引履歴获取实际卖出价格
                                actual_sell_price = None
                                if hasattr(self, 'trading_df') and self.trading_df is not None:
                                    try:
                                        # 查找对应日期的交易价格
                                        transaction_date = row['UTC_Time'].strftime('%Y-%m-%d %H:%M')
                                        actual_sell_price = self.get_price_for_date(transaction_date, coin)
                                        if actual_sell_price and actual_sell_price > 0:
                                            print(f"  ✓ 現物注文取引履歴から{coin}売却単価を取得: {actual_sell_price:,.0f} 円")
                                            # 使用实际价格重新计算卖出金额
                                            jpy_revenue_amount = quantity * actual_sell_price
                                            print(f"  ✓ 使用实际价格重新计算: {quantity:.8f} × {actual_sell_price:,.0f} = {jpy_revenue_amount:,.0f} 円")
                                        else:
                                            print(f"  ⚠️ 現物注文取引履歴から{coin}的価格を取得できませんでした")
                                    except Exception as e:
                                        print(f"  現物注文取引履歴からの価格取得エラー: {str(e)}")
                                
                                # 損益計算（手数料を考慮）
                                if method == "total_average":
                                    # 总平均法：使用总平均成本计算
                                    cost_basis = total_avg_cost * quantity
                                    print(f"  総平均法: 売却数量={quantity:.8f}, 総平均単価={total_avg_cost:,.0f} 円, 取得費={cost_basis:,.0f} 円")
                                else:
                                    # 移动平均法：使用当前平均成本计算
                                    cost_basis = avg_cost * quantity
                                    print(f"  移動平均法: 売却数量={quantity:.8f}, 現在平均単価={avg_cost:,.0f} 円, 取得費={cost_basis:,.0f} 円")
                                
                                profit = jpy_revenue_amount - cost_basis - jpy_fee_amount
                                
                                # 公式詳細を記録
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
                                
                                # 更详细的卖出信息显示
                                if method == "total_average":
                                    print(f"    【総平均法計算詳細】")
                                    print(f"    売却数量: {quantity:.8f} BTC")
                                    print(f"    売却代金: {jpy_revenue_amount:,.0f} 円")
                                    print(f"    総平均単価: {total_avg_cost:,.0f} 円/BTC")
                                    print(f"    売却成本: {cost_basis:,.0f} 円 (={quantity:.8f} × {total_avg_cost:,.0f})")
                                    print(f"    手数料: {jpy_fee_amount:,.0f} 円")
                                    print(f"    損益計算: {jpy_revenue_amount:,.0f} - {cost_basis:,.0f} - {jpy_fee_amount:,.0f} = {profit:,.0f} 円")
                                else:
                                    print(f"    【移動平均法計算詳細】")
                                    print(f"    売却数量: {quantity:.8f} BTC")
                                    print(f"    売却代金: {jpy_revenue_amount:,.0f} 円")
                                    print(f"    現在平均単価: {avg_cost:,.0f} 円/BTC")
                                    print(f"    売却成本: {cost_basis:,.0f} 円 (={quantity:.8f} × {avg_cost:,.0f})")
                                    print(f"    手数料: {jpy_fee_amount:,.0f} 円")
                                    print(f"    損益計算: {jpy_revenue_amount:,.0f} - {cost_basis:,.0f} - {jpy_fee_amount:,.0f} = {profit:,.0f} 円")
                                
                                # 更新持有量
                                holdings['BTC']['quantity'] -= quantity
                                holdings['BTC']['total_cost'] -= cost_basis
                                
                                # 売却単価を計算
                                sell_unit_price = "N/A"
                                if actual_sell_price and actual_sell_price > 0:
                                    sell_unit_price = f"{actual_sell_price:,.0f}"
                                elif quantity > 0:
                                    # 如果没有实际价格，使用计算出的单价
                                    calculated_price = jpy_revenue_amount / quantity
                                    sell_unit_price = f"{calculated_price:,.0f}"
                                
                                results.append([transaction_counter, date, "SELL", coin, quantity, sell_unit_price, f"{profit:,.0f}", f"{avg_cost:,.0f}"])
                                
                                # 処理済みとして記録 - 立即标记所有相关交易为已处理
                                processed_transactions.add(transaction_key)
                                for _, related in jpy_revenue.iterrows():
                                    related_key = f"{date}_{related['Operation']}_{related['Coin']}_{related['Change']}"
                                    processed_transactions.add(related_key)
                                    print(f"    関連取引を処理済みとして記録: {related_key}")
                                for _, related in jpy_fees.iterrows():
                                    related_key = f"{date}_{related['Operation']}_{related['Coin']}_{related['Change']}"
                                    processed_transactions.add(related_key)
                                    print(f"    関連取引を処理済みとして記録: {related_key}")
                                
                                print(f"    ✓ SELL取引完了: {transaction_key}")
                        else:
                            print(f"警告: BTCの保有数量が不足しています。必要: {quantity}, 保有: {holdings['BTC']['quantity']}")
            
            # 显示结果
            if results:
                self.result_df = pd.DataFrame(results, columns=["番号", "日付", "種別", "通貨", "数量", "単価(JPY)", "損益(JPY)", "平均取得単価"])
                
                # 计算总损益
                total_profit = sum([float(str(r[6]).replace(',', '')) for r in results if r[6] != 0 and str(r[6]) != "N/A"])
                
                # 计算买卖汇总信息
                total_buy_quantity = 0.0
                total_buy_amount = 0.0
                total_sell_quantity = 0.0
                total_sell_amount = 0.0
                
                print(f"\n=== 详细金额计算 ===")
                
                for result in results:
                    if result[2] == "BUY":  # 种别列
                        quantity = float(str(result[4]).replace(',', ''))  # 数量列
                        unit_price = str(result[5])  # 単価(JPY)列
                        
                        total_buy_quantity += quantity
                        
                        if unit_price != "N/A":
                            try:
                                amount = float(unit_price.replace(',', ''))
                                # 修复：计算总金额，而不是只加单价
                                total_amount = quantity * amount
                                total_buy_amount += total_amount
                                print(f"  BUY: {quantity:.8f} BTC × {amount:,.0f} 円 = {total_amount:,.0f} 円")
                            except:
                                print(f"  BUY: {quantity:.8f} BTC × {unit_price} (无法解析)")
                        else:
                            print(f"  BUY: {quantity:.8f} BTC × N/A (单价缺失)")
                            
                    elif result[2] == "SELL":  # 种别列
                        quantity = float(str(result[4]).replace(',', ''))  # 数量列
                        total_sell_quantity += quantity
                        
                        # 对于SELL，我们需要从損益(JPY)列计算金额
                        if result[6] != "N/A":
                            try:
                                profit = float(str(result[6]).replace(',', ''))
                                avg_cost = float(str(result[7]).replace(',', ''))
                                
                                # 修复：使用真实的卖出价格计算卖出金额
                                # 从単価(JPY)列获取真实价格
                                unit_price_str = str(result[5])
                                if unit_price_str != "N/A":
                                    try:
                                        real_unit_price = float(unit_price_str.replace(',', ''))
                                        
                                        # 添加价格合理性检查
                                        if real_unit_price > 50000000:  # 如果单价超过5000万日元，认为异常
                                            print(f"  ⚠️ 检测到异常高单价: {real_unit_price:,.0f} 円，使用回退方法")
                                            sell_amount = profit + (quantity * avg_cost)
                                            print(f"  SELL: {quantity:.8f} BTC, 損益: {profit:,.0f} 円, 取得費: {quantity * avg_cost:,.0f} 円, 売却金額: {sell_amount:,.0f} 円 (回退方法)")
                                        else:
                                            sell_amount = quantity * real_unit_price
                                            print(f"  SELL: {quantity:.8f} BTC, 損益: {profit:,.0f} 円, 取得費: {quantity * avg_cost:,.0f} 円, 売却金額: {sell_amount:,.0f} 円 (単価: {real_unit_price:,.0f} 円)")
                                    except:
                                        # 如果无法解析单价，回退到原来的方法
                                        sell_amount = profit + (quantity * avg_cost)
                                        print(f"  SELL: {quantity:.8f} BTC, 損益: {profit:,.0f} 円, 取得費: {quantity * avg_cost:,.0f} 円, 売却金額: {sell_amount:,.0f} 円 (回退方法)")
                                else:
                                    # 如果单价是N/A，回退到原来的方法
                                    sell_amount = profit + (quantity * avg_cost)
                                    print(f"  SELL: {quantity:.8f} BTC, 損益: {profit:,.0f} 円, 取得費: {quantity * avg_cost:,.0f} 円, 売却金額: {sell_amount:,.0f} 円 (回退方法)")
                                
                                total_sell_amount += sell_amount
                            except:
                                print(f"  SELL: {quantity:.8f} BTC (无法计算売却金額)")
                        else:
                            print(f"  SELL: {quantity:.8f} BTC (損益N/A)")
                
                print(f"\n=== 汇总计算结果 ===")
                print(f"总买入数量: {total_buy_quantity:.8f} BTC")
                print(f"总买入金额: {total_buy_amount:,.0f} 円")
                print(f"总卖出数量: {total_sell_quantity:.8f} BTC")
                print(f"总卖出金额: {total_sell_amount:,.0f} 円")
                if total_buy_quantity > 0:
                    print(f"平均买入单价: {total_buy_amount / total_buy_quantity:,.0f} 円/BTC")
                
                # 显示公式和结果
                self.display_formulas(formula_details, total_profit, results)
                
                print(f"\n=== 日元交易計算結果サマリー ===")
                print(f"総取引件数: {len(results)}")
                print(f"総買入 BTC: {total_buy_quantity:.8f}, 総価格: {total_buy_amount:,.0f} 円")
                print(f"総売出 BTC: {total_sell_quantity:.8f}, 総価格: {total_sell_amount:,.0f} 円")
                print(f"年間損益合計: {total_profit:,.0f} 円")
                print(f"通貨別保有状況:")
                for coin, data in holdings.items():
                    if data['quantity'] > 0:
                        print(f"  {coin}: 数量={data['quantity']}, 取得費={data['total_cost']:,.0f} 円")
                
                # 显示结果
                self.display_results()
                
                # 更新汇总信息显示 - 使用更详细的格式
                avg_buy_price = total_buy_amount / total_buy_quantity if total_buy_quantity > 0 else 0
                avg_sell_price = total_sell_amount / total_sell_quantity if total_sell_quantity > 0 else 0
                
                # 获取计算方法
                method = self.calculation_method.get()
                method_name = "総平均法" if method == "total_average" else "移動平均法"
                
                summary_text = f"計算方法: {method_name}\n"
                summary_text += f"総買入 BTC: {total_buy_quantity:.8f}, 単価: {avg_buy_price:,.0f} 円, 花费: {total_buy_amount:,.0f} 円\n"
                summary_text += f"総売出 BTC: {total_sell_quantity:.8f}, 単価: {avg_sell_price:,.0f} 円, 获得: {total_sell_amount:,.0f} 円\n"
                
                # 添加详细计算说明
                if method == "total_average":
                    total_avg_cost = total_buy_amount / total_buy_quantity if total_buy_quantity > 0 else 0
                    total_sell_cost = total_sell_quantity * total_avg_cost
                    total_fees = total_sell_amount - total_sell_cost - total_profit
                    summary_text += f"\n【総平均法計算詳細】\n"
                    summary_text += f"総平均単価: {total_avg_cost:,.0f} 円/BTC\n"
                    summary_text += f"総売出成本: {total_sell_cost:,.0f} 円 (={total_sell_quantity:.8f} × {total_avg_cost:,.0f})\n"
                    summary_text += f"総手数料: {total_fees:,.0f} 円\n"
                    summary_text += f"損益計算: {total_sell_amount:,.0f} - {total_sell_cost:,.0f} - {total_fees:,.0f} = {total_profit:,.0f} 円\n"
                else:
                    # 移动平均法的详细计算说明
                    summary_text += f"\n【移動平均法計算詳細】\n"
                    summary_text += f"移動平均法は各取引ごとに平均取得単価を再計算します\n"
                    summary_text += f"各売却取引の損益 = 売却代金 - (売却数量 × 当時の平均取得単価) - 手数料\n"
                    summary_text += f"年間損益合計 = 全売却取引の損益合計\n"
                
                summary_text += f"\n年間損益合計: {total_profit:,.0f} 円"
                
                self.total_profit_var.set(summary_text)
                
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
                    
                    # 获取所有需要的货币列表
                    all_coins = ['ETH', 'BTC', 'MATIC', 'ADA', 'DOT', 'LINK', 'SOL', 'XRP', 'AVAX', 'UNI', 'ATOM', 'LTC', 'BCH', 'ETC', 'FIL', 'NEAR', 'ALGO', 'VET', 'THETA', 'TRX', 'EOS', 'XLM']
                    
                    for i, date in enumerate(dates_list):
                        # 更新状态
                        progress_window.after(0, lambda d=date, idx=i: status_label.config(text=f"正在获取 {d} 的价格... ({idx+1}/{len(dates_list)})"))
                        
                        date_prices = {}
                        
                        # 获取ETH价格
                        eth_price = self.fetch_eth_price_for_date(date)
                        if eth_price > 0:
                            date_prices['ETH'] = eth_price
                        
                        # 获取BTC价格
                        btc_price = self.fetch_btc_price_for_date(date)
                        if btc_price > 0:
                            date_prices['BTC'] = btc_price
                        
                        # 获取其他货币价格（只获取前几个主要货币，避免API限制）
                        other_coins = ['MATIC', 'ADA', 'DOT', 'LINK', 'SOL']
                        for coin in other_coins:
                            coin_price = self.fetch_other_coin_price(coin, date)
                            if coin_price > 0:
                                date_prices[coin] = coin_price
                        
                        if date_prices:
                            prices_data[date] = date_prices
                        
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
    
    def fetch_other_coin_price(self, coin, date_str):
        """获取其他货币的价格（日元）"""
        try:
            import requests
            from datetime import datetime
            
            # 解析日期
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            timestamp = int(date_obj.timestamp())
            
            # 使用CoinGecko API获取指定货币的价格
            # 首先需要获取货币的ID
            coin_id = self.get_coin_id(coin)
            if not coin_id:
                print(f"  ✗ 无法找到{coin}的CoinGecko ID")
                return 0
            
            url = f"https://api.coingecko.com/api/v3/coins/{coin_id}/market_chart/range"
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
                    print(f"  ✓ {date_str} {coin}价格: {price:,.0f}円")
                    return price
            
            print(f"  ✗ {date_str} {coin}价格获取失败")
            return 0
            
        except Exception as e:
            print(f"  ✗ {date_str} {coin}价格获取错误: {e}")
            return 0
    
    def get_coin_id(self, coin):
        """获取货币的CoinGecko ID"""
        # 常见货币的ID映射
        coin_mapping = {
            'ADA': 'cardano',
            'DOT': 'polkadot',
            'LINK': 'chainlink',
            'MATIC': 'matic-network',
            'SOL': 'solana',
            'XRP': 'ripple',
            'AVAX': 'avalanche-2',
            'UNI': 'uniswap',
            'ATOM': 'cosmos',
            'LTC': 'litecoin',
            'BCH': 'bitcoin-cash',
            'ETC': 'ethereum-classic',
            'FIL': 'filecoin',
            'NEAR': 'near',
            'ALGO': 'algorand',
            'VET': 'vechain',
            'THETA': 'theta-token',
            'TRX': 'tron',
            'EOS': 'eos',
            'XLM': 'stellar'
        }
        
        return coin_mapping.get(coin.upper(), None)
    
    def show_price_result(self, prices_data, filename, window):
        """显示价格获取结果"""
        window.destroy()
        
        messagebox.showinfo("成功", f"历史价格获取完成！\n\n共获取了 {len(prices_data)} 个日期的价格\n已保存到: {filename}\n\n现在可以使用'损益计算执行'按钮进行更准确的计算！")
        
        print(f"\n=== 历史价格获取完成 ===")
        print(f"保存文件: {filename}")
        print("获取的价格数据:")
        for date, prices in prices_data.items():
            price_info = []
            for coin, price in prices.items():
                if price > 0:
                    price_info.append(f"{coin}={price:,.0f}円")
            print(f"  {date}: {', '.join(price_info)}")
    
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
        """根据日期和币种获取价格，优先使用現物注文取引履歴中的实际价格，支持交叉交易价格计算"""
        print(f"  🔍 开始获取{coin}价格，日期: {date_str}")
        
        # 首先尝试从現物注文取引履歴中获取实际价格
        if hasattr(self, 'trading_prices') and coin in self.trading_prices:
            print(f"    ✓ 找到trading_prices中的{coin}数据")
            # 将日期字符串转换为datetime对象
            try:
                date_obj = pd.to_datetime(date_str)
                # 查找最接近的日期
                min_diff = float('inf')
                best_price = None
                best_trade_info = None
                
                # 检查买入和卖出记录
                for trade_type in ['buys', 'sells']:
                    if trade_type in self.trading_prices[coin]:
                        for trade in self.trading_prices[coin][trade_type]:
                            # 严格过滤：只处理正确的交易对
                            pair = trade.get('pair', '')
                            if coin == 'BTC' and 'BTCJPY' not in pair:
                                print(f"      ⚠️ 跳过非BTCJPY交易对: {pair}")
                                continue  # 跳过非BTCJPY的交易
                            elif coin == 'ETH' and 'ETHJPY' not in pair:
                                print(f"      ⚠️ 跳过非ETHJPY交易对: {pair}")
                                continue  # 跳过非ETHJPY的交易
                            elif coin == 'DOGE' and 'DOGEJPY' not in pair:
                                print(f"      ⚠️ 跳过非DOGEJPY交易对: {pair}")
                                continue  # 跳过非DOGEJPY的交易
                            
                            diff = abs((trade['date'] - date_obj).total_seconds())
                            if diff < min_diff:
                                min_diff = diff
                                best_price = trade['price']
                                best_trade_info = f"{trade_type} {trade['date']} {trade['price']} ({pair})"
                
                if best_price is not None and min_diff <= 7200:  # 2小时内的价格
                    print(f"    ✓ 現物注文取引履歴から{coin}の価格を取得: {best_price} (日付差: {min_diff/60:.1f}分, 取引: {best_trade_info})")
                    return best_price
                elif best_price is not None:
                    print(f"    ⚠️ 現物注文取引履歴から{coin}の価格を取得: {best_price} (日付差: {min_diff/60:.1f}分, 取引: {best_trade_info}) - 日付差が大きすぎます")
                else:
                    print(f"    ❌ trading_prices中没有找到{coin}的正确价格数据（需要{coin}JPY交易对）")
            except Exception as e:
                print(f"    現物注文取引履歴からの価格取得エラー: {str(e)}")
        else:
            print(f"    ❌ 没有找到trading_prices中的{coin}数据")
        
        # 如果没有直接的币种价格，尝试通过交叉交易计算
        if hasattr(self, 'trading_df') and self.trading_df is not None:
            print(f"    🔍 尝试通过交叉交易计算{coin}价格")
            try:
                date_obj = pd.to_datetime(date_str)
                # 查找2小时内的相关交易
                time_window = pd.Timedelta(hours=2)
                
                # 尝试通过ETH/BTC交易计算BTC价格
                if coin == 'BTC':
                    print(f"      🔍 尝试ETH/BTC交叉交易计算")
                    # 查找ETH/BTC交易
                    eth_btc_trades = self.trading_df[
                        (self.trading_df['Pair'] == 'ETHBTC') & 
                        (pd.to_datetime(self.trading_df['Date(UTC)']) >= date_obj - time_window) &
                        (pd.to_datetime(self.trading_df['Date(UTC)']) <= date_obj + time_window)
                    ]
                    
                    if not eth_btc_trades.empty:
                        print(f"        ✓ 找到{len(eth_btc_trades)}个ETH/BTC交易")
                        # 使用最近的ETH/BTC交易价格
                        latest_trade = eth_btc_trades.iloc[-1]
                        eth_btc_price = float(latest_trade['Price'])
                        
                        # 尝试获取ETH的JPY价格
                        eth_jpy_trades = self.trading_df[
                            (self.trading_df['Pair'] == 'ETHJPY') & 
                            (pd.to_datetime(self.trading_df['Date(UTC)']) >= date_obj - time_window) &
                            (pd.to_datetime(self.trading_df['Date(UTC)']) <= date_obj + time_window)
                        ]
                        
                        if not eth_jpy_trades.empty:
                            print(f"        ✓ 找到{len(eth_jpy_trades)}个ETH/JPY交易")
                            eth_jpy_price = float(eth_jpy_trades.iloc[-1]['Price'])
                            # 计算BTC的JPY价格：BTC价格 = ETH价格 × ETH/BTC比率
                            btc_jpy_price = eth_jpy_price * eth_btc_price
                            print(f"        ✓ 交叉交易からBTC価格を計算: ETH/JPY {eth_jpy_price:,.0f} × ETH/BTC {eth_btc_price:.6f} = {btc_jpy_price:,.0f} 円")
                            return btc_jpy_price
                        else:
                            print(f"        ❌ 没有找到ETH/JPY交易")
                    else:
                        print(f"        ❌ 没有找到ETH/BTC交易")
                    
                    # 尝试从trading_prices中获取BTC价格
                    if hasattr(self, 'trading_prices') and 'BTC' in self.trading_prices:
                        print(f"      🔍 尝试从trading_prices获取BTC价格")
                        # 查找最接近的BTC交易
                        min_diff = float('inf')
                        best_price = None
                        
                        for trade_type in ['buys', 'sells']:
                            for trade in self.trading_prices['BTC'][trade_type]:
                                diff = abs((trade['date'] - date_obj).total_seconds())
                                if diff < min_diff:
                                    min_diff = diff
                                    # 如果是ETHBTC交易，需要计算JPY价格
                                    if 'eth_amount' in trade:
                                        # 尝试获取ETH的JPY价格
                                        eth_jpy_trades = self.trading_df[
                                            (self.trading_df['Pair'] == 'ETHJPY') & 
                                            (pd.to_datetime(self.trading_df['Date(UTC)']) >= trade['date'] - time_window) &
                                            (pd.to_datetime(self.trading_df['Date(UTC)']) <= trade['date'] + time_window)
                                        ]
                                        if not eth_jpy_trades.empty:
                                            eth_jpy_price = float(eth_jpy_trades.iloc[-1]['Price'])
                                            btc_jpy_price = eth_jpy_price * trade['price']  # trade['price']是BTC/ETH价格
                                            best_price = btc_jpy_price
                                            print(f"        ✓ 从ETHBTC交易计算BTC价格: {btc_jpy_price:,.0f} 円")
                                        else:
                                            # 如果没有ETHJPY价格，使用ETHBTC价格作为参考
                                            # 这里可以设置一个默认的ETH价格，或者跳过
                                            print(f"        ⚠️ 没有找到ETHJPY价格，无法计算BTC的JPY价格")
                        
                        if best_price is not None and min_diff <= 7200:  # 2小时内的价格
                            return best_price
                
                # 尝试通过DOGE/JPY交易计算（如果有的话）
                if coin == 'BTC':
                    print(f"      🔍 尝试DOGE/JPY交叉交易计算")
                    doge_jpy_trades = self.trading_df[
                        (self.trading_df['Pair'] == 'DOGEJPY') & 
                        (pd.to_datetime(self.trading_df['Date(UTC)']) >= date_obj - time_window) &
                        (pd.to_datetime(self.trading_df['Date(UTC)']) <= date_obj + time_window)
                    ]
                    
                    if not doge_jpy_trades.empty:
                        print(f"        ✓ 找到{len(doge_jpy_trades)}个DOGE/JPY交易")
                        # 查找DOGE/BTC交易
                        doge_btc_trades = self.trading_df[
                            (self.trading_df['Pair'] == 'DOGEBTC') & 
                            (pd.to_datetime(self.trading_df['Date(UTC)']) >= date_obj - time_window) &
                            (pd.to_datetime(self.trading_df['Date(UTC)']) <= date_obj + time_window)
                        ]
                        
                        if not doge_btc_trades.empty:
                            print(f"        ✓ 找到{len(doge_btc_trades)}个DOGE/BTC交易")
                            doge_jpy_price = float(doge_jpy_trades.iloc[-1]['Price'])
                            doge_btc_price = float(doge_btc_trades.iloc[-1]['Price'])
                            # 计算BTC的JPY价格：BTC价格 = DOGE价格 × DOGE/BTC比率
                            btc_jpy_price = doge_jpy_price * doge_btc_price
                            print(f"        ✓ 交叉交易からBTC価格を計算: DOGE/JPY {doge_jpy_price:,.0f} × DOGE/BTC {doge_btc_price:.8f} = {btc_jpy_price:,.0f} 円")
                            return btc_jpy_price
                        else:
                            print(f"        ❌ 没有找到DOGE/BTC交易")
                    else:
                        print(f"        ❌ 没有找到DOGE/JPY交易")
                            
            except Exception as e:
                print(f"    交叉交易価格計算エラー: {str(e)}")
        
        # 如果没有現物注文取引履歴价格，则使用历史价格
        print(f"    🔍 尝试使用历史价格")
        prices = self.load_historical_prices()
        date_key = date_str[:10]  # 只取日期部分，去掉时间
        
        if date_key in prices and coin in prices[date_key]:
            print(f"    ✓ 历史价格から{coin}の価格を取得: {prices[date_key][coin]}")
            return prices[date_key][coin]
        
        print(f"    ❌ {coin}の価格が見つかりません: {date_str}")
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
            # 获取计算方法
            method = self.calculation_method.get()
            method_name = "総平均法" if method == "total_average" else "移動平均法"
            
            total_formula += f"年間損益合計 ({method_name}) = "
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
    
    def save_results_by_coin(self, results):
        """按货币分别保存结果到Excel文件"""
        try:
            import pandas as pd
            from collections import defaultdict
            import os
            
            # 创建结果文件夹
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            results_folder = f"暗号資産計算結果_{timestamp}"
            os.makedirs(results_folder, exist_ok=True)
            
            # 按货币分组结果
            coin_results = defaultdict(list)
            for result in results:
                coin = result[3]  # 货币列
                coin_results[coin].append(result)
            
            # 为每个货币创建Excel文件
            for coin, coin_data in coin_results.items():
                if coin_data:
                    # 创建DataFrame
                    df = pd.DataFrame(coin_data, columns=["番号", "日付", "種別", "通貨", "数量", "単価(JPY)", "損益(JPY)", "平均取得単価"])
                    
                    # 计算该货币的汇总信息
                    total_quantity = sum([float(str(r[4]).replace(',', '')) for r in coin_data if r[4] != "N/A" and str(r[4]) != "0"])
                    total_profit = sum([float(str(r[6]).replace(',', '')) for r in coin_data if r[6] != 0 and str(r[6]) != "N/A"])
                    
                    # 文件名
                    filename = f"{coin}_取引記録.xlsx"
                    filepath = os.path.join(results_folder, filename)
                    
                    # 保存到Excel
                    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                        # 交易记录表
                        df.to_excel(writer, sheet_name='取引記録', index=False)
                        
                        # 汇总表
                        summary_data = {
                            '項目': ['総取引件数', '総数量', '総損益(円)', '平均単価(円)'],
                            '値': [len(coin_data), total_quantity, total_profit, total_profit/total_quantity if total_quantity > 0 else 0]
                        }
                        summary_df = pd.DataFrame(summary_data)
                        summary_df.to_excel(writer, sheet_name='サマリー', index=False)
                    
                    print(f"✓ {coin}の結果を保存しました: {filename}")
            
            print(f"✓ すべての結果を保存しました: {results_folder}")
            return results_folder
            
        except Exception as e:
            print(f"通貨別保存エラー: {e}")
            return None
    
    def generate_tax_report(self, results):
        """生成税务申报用的总计算表格"""
        try:
            import pandas as pd
            from collections import defaultdict
            import os
            
            # 按货币分组并计算汇总数据
            coin_summary = defaultdict(lambda: {
                'purchase_quantity': 0,
                'purchase_amount': 0,
                'sale_quantity': 0,
                'sale_amount': 0,
                'profit': 0,
                'fees': 0
            })
            
            for result in results:
                coin = result[3]  # 货币
                operation = result[2]  # 操作类型
                quantity = float(str(result[4]).replace(',', '')) if result[4] != "N/A" else 0
                amount = float(str(result[6]).replace(',', '')) if result[6] != "N/A" else 0
                
                if operation == "Transaction Buy":
                    coin_summary[coin]['purchase_quantity'] += quantity
                    coin_summary[coin]['purchase_amount'] += abs(amount)
                elif operation == "Transaction Sold":
                    coin_summary[coin]['sale_quantity'] += quantity
                    coin_summary[coin]['sale_amount'] += abs(amount)
                    coin_summary[coin]['profit'] += amount
                elif operation == "Transaction Fee":
                    coin_summary[coin]['fees'] += abs(amount)
                elif operation == "Simple Earn Locked Rewards":
                    coin_summary[coin]['purchase_quantity'] += quantity
                    coin_summary[coin]['purchase_amount'] += 0  # 奖励通常成本为0
            
            # 创建税务申报表格
            tax_data = []
            total_purchase_quantity = 0
            total_purchase_amount = 0
            total_sale_quantity = 0
            total_sale_amount = 0
            total_profit = 0
            total_fees = 0
            
            for coin, data in coin_summary.items():
                if data['purchase_quantity'] > 0 or data['sale_quantity'] > 0:
                    tax_data.append({
                        '暗号資産の名称': coin,
                        '購入数量': f"{data['purchase_quantity']:.8f}",
                        '購入金額': f"{data['purchase_amount']:,.0f}",
                        '売却数量': f"{data['sale_quantity']:.8f}",
                        '売却金額': f"{data['sale_amount']:,.0f}",
                        '損益': f"{data['profit']:,.0f}",
                        '手数料': f"{data['fees']:,.0f}"
                    })
                    
                    total_purchase_quantity += data['purchase_quantity']
                    total_purchase_amount += data['purchase_amount']
                    total_sale_quantity += data['sale_quantity']
                    total_sale_amount += data['sale_amount']
                    total_profit += data['profit']
                    total_fees += data['fees']
            
            # 添加总计行
            tax_data.append({
                '暗号資産の名称': '合計',
                '購入数量': f"{total_purchase_quantity:.8f}",
                '購入金額': f"{total_purchase_amount:,.0f}",
                '売却数量': f"{total_sale_quantity:.8f}",
                '売却金額': f"{total_sale_amount:,.0f}",
                '損益': f"{total_profit:,.0f}",
                '手数料': f"{total_fees:,.0f}"
            })
            
            # 创建DataFrame并保存
            tax_df = pd.DataFrame(tax_data)
            filename = "暗号資産税務申告表.xlsx"
            
            # 获取结果文件夹路径（从save_results_by_coin方法返回）
            results_folder = getattr(self, '_results_folder', None)
            if results_folder:
                filepath = os.path.join(results_folder, filename)
            else:
                # 如果没有结果文件夹，创建默认文件夹
                timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
                results_folder = f"暗号資産計算結果_{timestamp}"
                os.makedirs(results_folder, exist_ok=True)
                filepath = os.path.join(results_folder, filename)
            
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                # 主要汇总表
                tax_df.to_excel(writer, sheet_name='通貨別取引サマリー', index=False)
                
                # 日本税务申报表格格式
                japan_tax_data = {
                    '項目': [
                        '収入金額',
                        '売却価額',
                        '売却原価',
                        '必要経費',
                        '手数料等',
                        '所得金額'
                    ],
                    '金額(円)': [
                        total_sale_amount,
                        total_sale_amount,
                        total_purchase_amount,
                        total_fees,
                        total_fees,
                        total_profit - total_fees
                    ]
                }
                japan_tax_df = pd.DataFrame(japan_tax_data)
                japan_tax_df.to_excel(writer, sheet_name='日本税務申告用', index=False)
                
                # 详细计算表
                detail_data = {
                    '計算項目': [
                        '年始残高',
                        '購入等',
                        '売却等',
                        '売却原価',
                        '末残高・翌年繰越'
                    ],
                    '数量': [
                        0,  # 年始残高
                        total_purchase_quantity,
                        total_sale_quantity,
                        total_sale_quantity,  # 假设全部卖出
                        total_purchase_quantity - total_sale_quantity
                    ],
                    '金額(円)': [
                        0,  # 年始残高
                        total_purchase_amount,
                        total_sale_amount,
                        total_purchase_amount * (total_sale_quantity / total_purchase_quantity) if total_purchase_quantity > 0 else 0,
                        total_purchase_amount * ((total_purchase_quantity - total_sale_quantity) / total_purchase_quantity) if total_purchase_quantity > 0 else 0
                    ]
                }
                detail_df = pd.DataFrame(detail_data)
                detail_df.to_excel(writer, sheet_name='詳細計算', index=False)
            
            print(f"✓ 税務申告用総合計算表を保存しました: {filename}")
            
        except Exception as e:
            print(f"税務申告表生成エラー: {e}")
    
    def regenerate_files(self):
        """重新生成结果文件"""
        if not hasattr(self, 'result_df') or self.result_df is None:
            messagebox.showwarning("警告", "先に損益計算を実行してください")
            return
        
        try:
            # 获取计算结果
            results = []
            for _, row in self.result_df.iterrows():
                results.append(list(row))
            
            if not results:
                messagebox.showwarning("警告", "計算結果がありません")
                return
            
            # 按货币分别保存结果
            results_folder = self.save_results_by_coin(results)
            
            # 生成总的税务申报表格
            if results_folder:
                self._results_folder = results_folder
            self.generate_tax_report(results)
            
            # 显示成功消息
            if results_folder:
                messagebox.showinfo("成功", f"ファイルを再生成しました！\n\n保存先: {results_folder}\n\nフォルダを開きますか？")
                
                # 询问是否打开文件夹
                if messagebox.askyesno("確認", "結果フォルダを開きますか？"):
                    import os
                    os.startfile(results_folder)
            else:
                messagebox.showinfo("成功", "ファイルを再生成しました！")
                
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル再生成中にエラーが発生しました:\n{str(e)}")
            print(f"ファイル再生成エラー: {e}")

def main():
    root = tk.Tk()
    app = CryptoCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
