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
        """
        algo_text.insert(tk.END, algorithm_description)
        algo_text.config(state=tk.DISABLED)
        
        # 計算実行ボタン
        calc_btn = ttk.Button(main_frame, text="損益計算実行", command=self.calculate_profit)
        calc_btn.grid(row=3, column=0, columnspan=4, pady=20)
        
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
            # Binanceのデータ形式に対応
            # 必要なカラム: UTC_Time, Operation, Coin, Change
            
            # 取引データのみを抽出（入出金、報酬などは除外）
            trade_operations = [
                'Buy Crypto With Fiat',      # 日本円で暗号資産購入
                'Transaction Buy',           # 暗号資産購入
                'Transaction Sold',          # 暗号資産売却
                'Transaction Spend',         # 暗号資産支払い
                'Transaction Revenue',       # 暗号資産受取
                'Transaction Fee'            # 手数料
            ]
            
            # 取引データのみをフィルタリング
            trade_df = self.df[self.df['Operation'].isin(trade_operations)].copy()
            
            if len(trade_df) == 0:
                messagebox.showerror("エラー", "取引データが見つかりません")
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
            
            for idx, row in trade_df.iterrows():
                operation = row['Operation']
                coin = row['Coin']
                change = float(row['Change'])
                date = row['UTC_Time'].strftime('%Y-%m-%d %H:%M')
                
                # 重複処理防止
                transaction_key = f"{date}_{operation}_{coin}_{change}"
                if transaction_key in processed_transactions:
                    print(f"重複取引をスキップ: {transaction_key}")
                    continue
                
                # デバッグ用：各取引の詳細を表示
                print(f"処理中: {date} | {operation} | {coin} | {change}")
                
                if coin not in holdings:
                    holdings[coin] = {'quantity': 0.0, 'total_cost': 0.0}
                
                if operation == 'Buy Crypto With Fiat':
                    # 日本円で暗号資産購入
                    if coin != 'JPY':
                        transaction_counter += 1  # 取引番号を増加
                        
                        # 暗号資産の数量を正の値に
                        quantity = abs(change)
                        old_quantity = holdings[coin]['quantity']
                        old_cost = holdings[coin]['total_cost']
                        
                        # 日本円の支出を取得費に加算
                        jpy_row = trade_df[(trade_df['UTC_Time'] == row['UTC_Time']) & (trade_df['Coin'] == 'JPY') & (trade_df['Operation'] == 'Buy Crypto With Fiat')]
                        if not jpy_row.empty:
                            jpy_amount = abs(float(jpy_row.iloc[0]['Change']))
                            
                            # 手数料を取得（Transaction Feeから）
                            fee_row = trade_df[(trade_df['UTC_Time'] == row['UTC_Time']) & (trade_df['Operation'] == 'Transaction Fee')]
                            fee_amount = 0.0
                            if not fee_row.empty:
                                fee_amount = abs(float(fee_row.iloc[0]['Change']))
                            
                            # 保有数量と取得費を更新
                            holdings[coin]['quantity'] += quantity
                            holdings[coin]['total_cost'] += jpy_amount + fee_amount
                            
                            # 公式詳細を記録
                            new_quantity = holdings[coin]['quantity']
                            new_cost = holdings[coin]['total_cost']
                            avg_cost = new_cost / new_quantity if new_quantity > 0 else 0
                            
                            formula_detail = f"""
【取引{transaction_counter}】{date} {coin}購入取引
購入数量: {quantity}
購入金額: {jpy_amount}円
手数料: {fee_amount}円
購入前保有数量: {old_quantity}
購入前累計取得費: {old_cost}円
購入後保有数量: {new_quantity}
購入後累計取得費: {new_cost}円
平均取得単価: {new_cost}円 ÷ {new_quantity} = {avg_cost:.2f}円
"""
                            formula_details.append(formula_detail)
                            
                            print(f"  → {coin}購入: 数量={quantity}, 金額={jpy_amount}円, 手数料={fee_amount}円, 累計取得費={holdings[coin]['total_cost']}円")
                            
                            results.append([transaction_counter, date, "BUY", coin, quantity, round(jpy_amount/quantity, 2), 0, round(avg_cost, 2)])
                            
                            # 処理済みとして記録
                            processed_transactions.add(transaction_key)
                            processed_transactions.add(f"{date}_Buy Crypto With Fiat_JPY_{jpy_amount}")
                            if fee_amount > 0:
                                processed_transactions.add(f"{date}_Transaction Fee_JPY_{fee_amount}")
                
                elif operation == 'Transaction Sold':
                    # 暗号資産売却
                    if change < 0:
                        transaction_counter += 1  # 取引番号を増加
                        
                        quantity = abs(change)
                        if holdings[coin]['quantity'] >= quantity:
                            # 平均取得単価で損益計算
                            old_quantity = holdings[coin]['quantity']
                            old_cost = holdings[coin]['total_cost']
                            avg_cost = old_cost / old_quantity if old_quantity > 0 else 0
                            cost_basis = avg_cost * quantity
                            
                            # 売却代金を取得（Transaction Revenueから）
                            revenue_row = trade_df[(trade_df['UTC_Time'] == row['UTC_Time']) & (trade_df['Operation'] == 'Transaction Revenue')]
                            if not revenue_row.empty:
                                proceeds = float(revenue_row.iloc[0]['Change'])
                                
                                # 手数料を取得（Transaction Feeから）
                                fee_row = trade_df[(trade_df['UTC_Time'] == row['UTC_Time']) & (trade_df['Operation'] == 'Transaction Fee')]
                                fee_amount = 0.0
                                if not fee_row.empty:
                                    fee_amount = abs(float(fee_row.iloc[0]['Change']))
                                
                                # 損益計算（手数料を考慮）
                                profit = proceeds - cost_basis - fee_amount
                                
                                # 公式詳細を記録
                                new_quantity = old_quantity - quantity
                                new_cost = old_cost - cost_basis
                                
                                formula_detail = f"""
【取引{transaction_counter}】{date} {coin}売却取引
売却数量: {quantity}
売却代金: {proceeds}円
手数料: {fee_amount}円
売却前保有数量: {old_quantity}
売却前累計取得費: {old_cost}円
平均取得単価: {old_cost}円 ÷ {old_quantity} = {avg_cost:.2f}円
取得費（売却分）: {avg_cost:.2f}円 × {quantity} = {cost_basis:.2f}円
損益計算: {proceeds}円 - {cost_basis:.2f}円 - {fee_amount}円 = {profit:.2f}円
売却後保有数量: {new_quantity}
売却後累計取得費: {new_cost:.2f}円
"""
                                formula_details.append(formula_detail)
                                
                                print(f"  → {coin}売却: 数量={quantity}, 売却代金={proceeds}円, 取得費={cost_basis}円, 手数料={fee_amount}円, 損益={profit}円")
        
        # 保有数量と取得費を減らす
                                holdings[coin]['quantity'] -= quantity
                                holdings[coin]['total_cost'] -= cost_basis
                                
                                results.append([transaction_counter, date, "SELL", coin, quantity, "N/A", round(profit, 2), round(avg_cost, 2)])
                                
                                # 処理済みとして記録
                                processed_transactions.add(transaction_key)
                                processed_transactions.add(f"{date}_Transaction Revenue_JPY_{proceeds}")
                                if fee_amount > 0:
                                    processed_transactions.add(f"{date}_Transaction Fee_JPY_{fee_amount}")
                        else:
                            print(f"警告: {coin}の保有数量が不足しています。必要: {quantity}, 保有: {holdings[coin]['quantity']}")

# 結果をDataFrameに
            if results:
                self.result_df = pd.DataFrame(results, columns=["番号", "日付", "種別", "通貨", "数量", "単価(JPY)", "損益(JPY)", "平均取得単価"])

# 年間合計損益
                total_profit = self.result_df["損益(JPY)"].sum()
                
                # 公式詳細を表示
                self.display_formulas(formula_details, total_profit, results)
                
                print(f"\n=== 計算結果サマリー ===")
                print(f"総取引件数: {len(results)}")
                print(f"年間損益合計: {total_profit:,.0f}円")
                print(f"通貨別保有状況:")
                for coin, data in holdings.items():
                    if data['quantity'] > 0:
                        print(f"  {coin}: 数量={data['quantity']}, 取得費={data['total_cost']}円")
                
                # 結果を表示
                self.display_results()
                self.total_profit_var.set(f"年間損益合計: {total_profit:,.0f} 円")
                
                messagebox.showinfo("完了", "損益計算が完了しました！")
            else:
                messagebox.showwarning("警告", "計算可能な取引データが見つかりませんでした")
            
        except Exception as e:
            messagebox.showerror("エラー", f"計算中にエラーが発生しました:\n{str(e)}")
            print(f"エラー詳細: {e}")
    
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
                profit = result[6]  # 損益(JPY)のインデックスを調整
                profit_terms.append(f"取引{transaction_num}の損益({profit}円)")
            
            total_formula += " + ".join(profit_terms)
            total_profit = sum([r[6] for r in sell_results])  # 損益を再計算
            total_formula += f" = {total_profit:,.0f}円"
        
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
