Attribute VB_Name = "Module1"

'oΧV[gΜoΝ{^πNbN΅½Μ
Sub oΧΛΜoΝ{^_Click()
    
'v^[ζΚπΔΡo·
    Worksheets("oΧΛ").PrintPreview EnableChanges:=Ture
    
End Sub
'oΧV[gΜoΝsIπ{^πNbN΅½Μ
Sub oΝsIπ_Click()

'ZvZXπρ\¦
With Application
     .Calculation = xlCalculationManual
     .EnableEvents = False
     .ScreenUpdating = False
End With

'oΧV[gΜsΖρΜΟπθ`
Dim r As Long
selected_row = Selection.Row  'oΧV[gΕIπ΅½ZΜsΜlπselected_rowΙγό

'ϊ»

Worksheets("oΧΛ").Range("AB1").Value = ""
Worksheets("oΧΛ").Range("AD1").Value = ""
Worksheets("oΧΛ").Range("A3").Value = ""
Worksheets("oΧΛ").Range("H19").Value = ""
Worksheets("oΧΛ").Range("H22").Value = ""
Worksheets("oΧΛ").Range("K22").Value = ""
Worksheets("oΧΛ").Range("Z22").Value = ""
Worksheets("oΧΛ").Range("I40").Value = ""

For i = 1 To 12
    Worksheets("oΧΛ").Range("I" & (32 + i)).Value = ""
Next i

For i = 1 To 4
    Worksheets("oΧΛ").Range("H" & (22 + i * 2)).Value = ""
    Worksheets("oΧΛ").Range("K" & (22 + i * 2)).Value = ""
    Worksheets("oΧΛ").Range("Z" & (22 + i * 2)).Value = ""
Next i

'oΧΛV[gΜeΪiΚάΕjΙoΧΛΜf[^πζΎ΅γό
Worksheets("oΧΛ").Range("AB1").Value = Worksheets("oΧ").Range("E" & selected_row).Value
Worksheets("oΧΛ").Range("AD1").Value = Worksheets("oΧ").Range("G" & selected_row).Value
Worksheets("oΧΛ").Range("A3").Value = Worksheets("oΧ").Range("B" & selected_row).Value
Worksheets("oΧΛ").Range("H19").Value = Worksheets("oΧ").Range("L" & selected_row).Value
Worksheets("oΧΛ").Range("H22").Value = Worksheets("oΧ").Range("F" & selected_row).Value
Worksheets("oΧΛ").Range("K22").Value = Worksheets("oΧ").Range("C" & selected_row).Value
Worksheets("oΧΛ").Range("Z22").Value = Worksheets("oΧ").Range("M" & selected_row).Value
Worksheets("oΧΛ").Range("I40").Value = Worksheets("oΧ").Range("N" & selected_row).Value

'oΧV[gΜXΦΤ`tζΌΜόΝf[^πΟΙγόipostcode`recipient)
postcode = Worksheets("oΧ").Range("I" & selected_row).Value
address1 = Worksheets("oΧ").Range("J" & selected_row).Value
address2 = Worksheets("oΧ").Range("K" & selected_row).Value

code = Worksheets("oΧ").Range("F" & selected_row).Value
recipient = Worksheets("oΧ").Range("H" & selected_row).Value
numbers = Worksheets("oΧ").Range("M" & selected_row).Value
 
Dim x
x = 1
Do While Worksheets("oΧ").Range("H" & (selected_row + x)).Value = "γ―" Or Worksheets("oΧ").Range("H" & (selected_row + x)).Value = ""
    Worksheets("oΧΛ").Range("H" & (22 + x * 2)).Value = Worksheets("oΧ").Range("F" & (selected_row + x)).Value
    Worksheets("oΧΛ").Range("K" & (22 + x * 2)).Value = Worksheets("oΧ").Range("C" & (selected_row + x)).Value
    Worksheets("oΧΛ").Range("Z" & (22 + x * 2)).Value = Worksheets("oΧ").Range("M" & (selected_row + x)).Value
    If x = 4 Then Exit Do
    x = x + 1
Loop




'XΦΤͺΘ’κΜoΧΛππͺςΙζθπͺ―ι
'XΦΤͺσΕΝΘ’κΜ
If IsEmpty(postcode) = False Then
Worksheets("oΧΛ").Range("I33").Value = "§" & postcode
Worksheets("oΧΛ").Range("I34").Value = address1
Worksheets("oΧΛ").Range("I35").Value = address2
Worksheets("oΧΛ").Range("I36").Value = recipient
'»ΜΌΜκΜiXΦΤͺσj
Else
Worksheets("oΧΛ").Range("I34").Value = recipient
Worksheets("oΧΛ").Range("I33").Value = ""
Worksheets("oΧΛ").Range("I36").Value = ""
Worksheets("oΧΛ").Range("I35").Value = ""
End If

'ZΚπ\¦
With Application
     .Calculation = xlCalculationAutomatic
     .EnableEvents = True
     .ScreenUpdating = True
End With

End Sub

