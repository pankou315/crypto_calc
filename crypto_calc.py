import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

class CryptoCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("暗号資産損益計算機")
        self.root.geometry("1000x700")
        
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
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # ファイル選択部分
        file_frame = ttk.LabelFrame(main_frame, text="ファイル選択", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=60)
        file_entry.grid(row=0, column=0, padx=(0, 10))
        
        browse_btn = ttk.Button(file_frame, text="ファイル選択", command=self.browse_file)
        browse_btn.grid(row=0, column=1)
        
        load_btn = ttk.Button(file_frame, text="データ読み込み", command=self.load_data)
        load_btn.grid(row=0, column=2, padx=(10, 0))
        
        # アルゴリズム説明部分
        algo_frame = ttk.LabelFrame(main_frame, text="計算アルゴリズム", padding="10")
        algo_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        algo_text = tk.Text(algo_frame, height=8, width=80, wrap=tk.WORD)
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
        calc_btn.grid(row=3, column=0, columnspan=3, pady=20)
        
        # 結果表示部分
        result_frame = ttk.LabelFrame(main_frame, text="計算結果", padding="10")
        result_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
        # 結果表示用のTreeview
        self.tree = ttk.Treeview(result_frame, columns=("日付", "種別", "通貨", "数量", "単価(JPY)", "損益(JPY)", "平均取得単価"), show="headings", height=10)
        
        # カラムの設定
        self.tree.heading("日付", text="日付")
        self.tree.heading("種別", text="種別")
        self.tree.heading("通貨", text="通貨")
        self.tree.heading("数量", text="数量")
        self.tree.heading("単価(JPY)", text="単価(JPY)")
        self.tree.heading("損益(JPY)", text="損益(JPY)")
        self.tree.heading("平均取得単価", text="平均取得単価")
        
        # カラム幅の設定
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
        
        # 年間損益合計表示
        self.total_profit_var = tk.StringVar()
        self.total_profit_var.set("年間損益合計: 計算してください")
        total_profit_label = ttk.Label(result_frame, textvariable=self.total_profit_var, font=("Arial", 12, "bold"))
        total_profit_label.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # グリッドの重み設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
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
            
            for idx, row in trade_df.iterrows():
                operation = row['Operation']
                coin = row['Coin']
                change = float(row['Change'])
                date = row['UTC_Time'].strftime('%Y-%m-%d %H:%M')
                
                # デバッグ用：各取引の詳細を表示
                print(f"処理中: {date} | {operation} | {coin} | {change}")
                
                if coin not in holdings:
                    holdings[coin] = {'quantity': 0.0, 'total_cost': 0.0}
                
                if operation == 'Buy Crypto With Fiat':
                    # 日本円で暗号資産購入
                    if coin != 'JPY':
                        # 暗号資産の数量を正の値に
                        quantity = abs(change)
                        holdings[coin]['quantity'] += quantity
                        # 日本円の支出を取得費に加算
                        jpy_row = trade_df[(trade_df['UTC_Time'] == row['UTC_Time']) & (trade_df['Coin'] == 'JPY') & (trade_df['Operation'] == 'Buy Crypto With Fiat')]
                        if not jpy_row.empty:
                            jpy_amount = abs(float(jpy_row.iloc[0]['Change']))
                            holdings[coin]['total_cost'] += jpy_amount
                            
                            print(f"  → {coin}購入: 数量={quantity}, 金額={jpy_amount}円, 累計取得費={holdings[coin]['total_cost']}円")
                            
                        avg_cost = holdings[coin]['total_cost'] / holdings[coin]['quantity'] if holdings[coin]['quantity'] > 0 else 0
                        results.append([date, "BUY", coin, quantity, round(jpy_amount/quantity, 2), 0, round(avg_cost, 2)])
                
                elif operation == 'Transaction Sold':
                    # 暗号資産売却
                    if change < 0:
                        quantity = abs(change)
                        if holdings[coin]['quantity'] >= quantity:
                            # 平均取得単価で損益計算
                            avg_cost = holdings[coin]['total_cost'] / holdings[coin]['quantity'] if holdings[coin]['quantity'] > 0 else 0
                            cost_basis = avg_cost * quantity
                            
                            # 売却代金を取得（Transaction Revenueから）
                            revenue_row = trade_df[(trade_df['UTC_Time'] == row['UTC_Time']) & (trade_df['Operation'] == 'Transaction Revenue')]
                            if not revenue_row.empty:
                                proceeds = float(revenue_row.iloc[0]['Change'])
                                profit = proceeds - cost_basis
                                
                                print(f"  → {coin}売却: 数量={quantity}, 売却代金={proceeds}円, 取得費={cost_basis}円, 損益={profit}円")
        
        # 保有数量と取得費を減らす
                                holdings[coin]['quantity'] -= quantity
                                holdings[coin]['total_cost'] -= cost_basis
                                
                                results.append([date, "SELL", coin, quantity, "N/A", round(profit, 2), round(avg_cost, 2)])

# 結果をDataFrameに
            if results:
                self.result_df = pd.DataFrame(results, columns=["日付", "種別", "通貨", "数量", "単価(JPY)", "損益(JPY)", "平均取得単価"])

# 年間合計損益
                total_profit = self.result_df["損益(JPY)"].sum()
                
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
