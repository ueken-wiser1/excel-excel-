import numpy as np
from scipy.optimize import minimize
import pandas as pd

# Excelファイルからデータを読み込む
df = pd.read_excel('C:/Users/touko/OneDrive/株価分析/excel/株式データ/test/20230516/20230101_OSCI.xlsx')

# 3変数の値を含む列を抽出してnumpy配列に変換
rankings = df[['RSIスコア', 'ボリンジャーバンドスコア', 'MACDスコア']].values

# 疑似スコアを含む列を抽出してnumpy配列に変換
pseudo_scores = df['PseudoScore'].values

# 多項式
def f(coefficients, x):
    a, b, c = coefficients
    return a*x[0] + b*x[1] + c*x[2]

# 最小二乗誤差
def objective_function(coefficients):
    return np.sum((pseudo_scores - f(coefficients, rankings))**2)

# 初期値（a,b,c）
initial_guess = [1, 1, 1]

# 最適化
result = minimize(objective_function, initial_guess)

# 結果
print(f'Optimized coefficients are {result.x}')