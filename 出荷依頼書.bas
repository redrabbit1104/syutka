Attribute VB_Name = "Module1"

'出荷シートの出力ボタンをクリックした時の処理
Sub 出荷依頼書の出力ボタン_Click()
    
'プリンター画面を呼び出す
    Worksheets("出荷依頼書").PrintPreview EnableChanges:=Ture
    
End Sub
'出荷シートの出力行選択ボタンをクリックした時の処理
Sub 出力行選択_Click()

'演算プロセスを非表示
With Application
     .Calculation = xlCalculationManual
     .EnableEvents = False
     .ScreenUpdating = False
End With

'出荷シートの行と列の変数を定義
Dim r As Long
selected_row = Selection.Row  '出荷シートで選択したセルの行の値をselected_rowに代入

'初期化

Worksheets("出荷依頼書").Range("AB1").Value = ""
Worksheets("出荷依頼書").Range("AD1").Value = ""
Worksheets("出荷依頼書").Range("A3").Value = ""
Worksheets("出荷依頼書").Range("H19").Value = ""
Worksheets("出荷依頼書").Range("H22").Value = ""
Worksheets("出荷依頼書").Range("K22").Value = ""
Worksheets("出荷依頼書").Range("Z22").Value = ""
Worksheets("出荷依頼書").Range("I40").Value = ""

For i = 1 To 12
    Worksheets("出荷依頼書").Range("I" & (32 + i)).Value = ""
Next i

For i = 1 To 4
    Worksheets("出荷依頼書").Range("H" & (22 + i * 2)).Value = ""
    Worksheets("出荷依頼書").Range("K" & (22 + i * 2)).Value = ""
    Worksheets("出荷依頼書").Range("Z" & (22 + i * 2)).Value = ""
Next i

'出荷依頼書シートの各項目（数量まで）に出荷依頼書のデータを取得し代入
Worksheets("出荷依頼書").Range("AB1").Value = Worksheets("出荷").Range("E" & selected_row).Value
Worksheets("出荷依頼書").Range("AD1").Value = Worksheets("出荷").Range("G" & selected_row).Value
Worksheets("出荷依頼書").Range("A3").Value = Worksheets("出荷").Range("B" & selected_row).Value
Worksheets("出荷依頼書").Range("H19").Value = Worksheets("出荷").Range("L" & selected_row).Value
Worksheets("出荷依頼書").Range("H22").Value = Worksheets("出荷").Range("F" & selected_row).Value
Worksheets("出荷依頼書").Range("K22").Value = Worksheets("出荷").Range("C" & selected_row).Value
Worksheets("出荷依頼書").Range("Z22").Value = Worksheets("出荷").Range("M" & selected_row).Value
Worksheets("出荷依頼書").Range("I40").Value = Worksheets("出荷").Range("N" & selected_row).Value

'出荷シートの郵便番号～送付先名の入力データを変数に代入（postcode～recipient)
postcode = Worksheets("出荷").Range("I" & selected_row).Value
address1 = Worksheets("出荷").Range("J" & selected_row).Value
address2 = Worksheets("出荷").Range("K" & selected_row).Value

code = Worksheets("出荷").Range("F" & selected_row).Value
recipient = Worksheets("出荷").Range("H" & selected_row).Value
numbers = Worksheets("出荷").Range("M" & selected_row).Value
 
Dim x
x = 1
Do While Worksheets("出荷").Range("H" & (selected_row + x)).Value = "上同" Or Worksheets("出荷").Range("H" & (selected_row + x)).Value = ""
    Worksheets("出荷依頼書").Range("H" & (22 + x * 2)).Value = Worksheets("出荷").Range("F" & (selected_row + x)).Value
    Worksheets("出荷依頼書").Range("K" & (22 + x * 2)).Value = Worksheets("出荷").Range("C" & (selected_row + x)).Value
    Worksheets("出荷依頼書").Range("Z" & (22 + x * 2)).Value = Worksheets("出荷").Range("M" & (selected_row + x)).Value
    If x = 4 Then Exit Do
    x = x + 1
Loop




'郵便番号がない場合の出荷依頼書を条件分岐により処理を分ける
'郵便番号が空ではない場合の処理
If IsEmpty(postcode) = False Then
Worksheets("出荷依頼書").Range("I33").Value = "〒" & postcode
Worksheets("出荷依頼書").Range("I34").Value = address1
Worksheets("出荷依頼書").Range("I35").Value = address2
Worksheets("出荷依頼書").Range("I36").Value = recipient
'その他の場合の処理（郵便番号が空）
Else
Worksheets("出荷依頼書").Range("I34").Value = recipient
Worksheets("出荷依頼書").Range("I33").Value = ""
Worksheets("出荷依頼書").Range("I36").Value = ""
Worksheets("出荷依頼書").Range("I35").Value = ""
End If

'演算結果を表示
With Application
     .Calculation = xlCalculationAutomatic
     .EnableEvents = True
     .ScreenUpdating = True
End With

End Sub

