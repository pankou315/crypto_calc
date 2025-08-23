#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日期匹配和价格获取测试脚本
"""

import pandas as pd
import os
from datetime import datetime, timedelta

def test_date_matching():
    """测试两个文件的日期匹配"""
    
    # 文件路径
    transaction_file = "取りれ歴トランザクション記録のエクスポート.csv"
    trading_file = "現物注文取引履歴.csv"
    
    if not os.path.exists(transaction_file):
        print(f"❌ 文件不存在: {transaction_file}")
        return
        
    if not os.path.exists(trading_file):
        print(f"❌ 文件不存在: {trading_file}")
        return
    
    try:
        # 读取交易记录文件
        print("=== 读取交易记录文件 ===")
        df_transaction = pd.read_csv(transaction_file, encoding='utf-8')
        print(f"交易记录行数: {len(df_transaction)}")
        
        # 筛选SELL交易
        sell_transactions = df_transaction[df_transaction['Operation'] == 'Transaction Sold']
        print(f"SELL交易数量: {len(sell_transactions)}")
        
        # 显示前几个SELL交易
        print("\n前5个SELL交易:")
        for idx, row in sell_transactions.head().iterrows():
            date = row['UTC_Time']
            coin = row['Coin']
            change = row['Change']
            print(f"  {date} | {coin} | {change}")
        
        # 读取現物注文取引履歴文件
        print("\n=== 读取現物注文取引履歴文件 ===")
        df_trading = pd.read_csv(trading_file, encoding='utf-8')
        print(f"現物注文取引履歴行数: {len(df_trading)}")
        
        # 转换日期格式
        df_trading['Date(UTC)'] = pd.to_datetime(df_trading['Date(UTC)'])
        
        # 显示BTC相关的交易
        btc_trades = df_trading[df_trading['Pair'].str.contains('BTC')]
        print(f"BTC相关交易数量: {len(btc_trades)}")
        
        print("\n前5个BTC交易:")
        for idx, row in btc_trades.head().iterrows():
            date = row['Date(UTC)']
            pair = row['Pair']
            side = row['Side']
            price = row['Price']
            executed = row['Executed']
            print(f"  {date} | {pair} | {side} | {price} | {executed}")
        
        # 测试日期匹配
        print("\n=== 测试日期匹配 ===")
        
        # 选择一个SELL交易进行测试
        if len(sell_transactions) > 0:
            test_sell = sell_transactions.iloc[0]
            test_date = pd.to_datetime(test_sell['UTC_Time'])
            test_coin = test_sell['Coin']
            
            print(f"测试SELL交易: {test_date} | {test_coin}")
            
            # 在現物注文取引履歴中查找相近日期的交易
            min_diff = float('inf')
            best_match = None
            
            for idx, row in btc_trades.iterrows():
                trade_date = row['Date(UTC)']
                diff = abs((trade_date - test_date).total_seconds())
                
                if diff < min_diff:
                    min_diff = diff
                    best_match = row
            
            if best_match is not None:
                print(f"最匹配的交易:")
                print(f"  日期: {best_match['Date(UTC)']}")
                print(f"  交易对: {best_match['Pair']}")
                print(f"  方向: {best_match['Side']}")
                print(f"  价格: {best_match['Price']}")
                print(f"  执行: {best_match['Executed']}")
                print(f"  时间差: {min_diff/60:.1f}分钟")
                
                if min_diff <= 7200:  # 2小时内
                    print("  ✅ 时间差在2小时内，可以匹配")
                else:
                    print("  ⚠️ 时间差超过2小时，无法匹配")
            else:
                print("❌ 没有找到匹配的交易")
        
    except Exception as e:
        print(f"❌ 测试过程中发生错误: {str(e)}")

if __name__ == "__main__":
    test_date_matching()
