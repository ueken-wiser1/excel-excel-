Dim objXLS
 
 Set objXLS= WScript.CreateObject("Excel.Application")
 objXLS.Visible = True
 
'
'*------------------------------------------------------------------
  
 objXLS.Workbooks.Open("C:\Users\touko\OneDrive\��������\excel\�����f�[�^\allkabu1.xlsm") '�t�@�C���̏ꏊ
 
 objXLS.Application.Run "Module1.�s�v�^�O�폜test" '�N�����郂�W���[�����ƃ}�N���̖��O
 
 
'*------------------------------------------------------------------
 objXLS.Application.Quit
 
 Set obj = Nothing