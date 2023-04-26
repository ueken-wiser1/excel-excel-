Dim objXLS
 
 Set objXLS= WScript.CreateObject("Excel.Application")
 objXLS.Visible = True
 
'
'*------------------------------------------------------------------
  
 objXLS.Workbooks.Open("C:\Users\touko\OneDrive\株価分析\excel\株式データ\allkabu1.xlsm") 'ファイルの場所
 
 objXLS.Application.Run "Module1.不要タグ削除test" '起動するモジュール名とマクロの名前
 
 
'*------------------------------------------------------------------
 objXLS.Application.Quit
 
 Set obj = Nothing