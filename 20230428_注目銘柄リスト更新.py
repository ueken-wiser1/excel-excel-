#対象ファイル：注目銘柄
#最終列だけを見ていくコード
#最終列の左側に数字がなければ、その行はスキップ
#最終列の左側に数字があった場合、その列の証券コードを読み込み
#当日の日次データで同じ証券コードをサーチ
#あったら、その行の終値を最終列に書込
#書き込んだ終値が左側の数字より大きければオレンジ塗り、小さければ青塗り
#最終列の数字とその行の一番左側にある数字の差を銘柄名の右側に書込
#→注目してからその銘柄は上がっているのか下がっているのか見極め
#その数字は注目した時の終値の何%かを書込
#注目してからその数字が2%を超えるまで何日かかったか、日付の差分の書込
