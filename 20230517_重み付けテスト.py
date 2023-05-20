import numpy as np
from scipy.optimize import minimize
import pandas as pd

# Excelファイルからデータを読み込む
df = pd.read_excel('C:/Users/touko/OneDrive/株価分析/excel/株式データ/test/20230516/20230216_OSCI.xlsx')

# 3変数の値を含む列を抽出してnumpy配列に変換
rankings = df[['RSIスコア', 'ボリンジャーバンドスコア', 'MACDスコア']].values

# 疑似スコアを含む列を抽出してnumpy配列に変換
pseudo_scores = df['項.1'].values

# 多項式
def f(coefficients, x):
    return np.dot(x, coefficients)

# 最小二乗誤差
def objective_function(coefficients):
    return np.sum((pseudo_scores - f(coefficients, rankings))**2)

# 初期値（a,b,c）
initial_guess = [1, 1, 1]

# 最適化
result = minimize(objective_function, initial_guess)

# 結果を個別の変数に代入
a_optimized, b_optimized, c_optimized = result.x

# 結果を出力
print(f'Optimized coefficients are a={a_optimized}, b={b_optimized}, c={c_optimized}')
