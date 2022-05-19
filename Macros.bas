Attribute VB_Name = "Module1"
Option Explicit

Sub НоминалКупона()
Attribute НоминалКупона.VB_ProcData.VB_Invoke_Func = "П\n14"


Dim celF As Range, celHAN As Range, celNEW As Range, intHAN As Double, intNEW As Double, cHAN As Double, cHN As Double, cNEW As Double, cNW As Double, NumWorkSh As Integer, NumWS As Integer
NumWS = 1
NumWorkSh = 2
Set celF = Application.InputBox("Выберите столбец с категорией", Type:=8)
For Each celF In celF
    celF.FormulaR1C1 = "=RANDBETWEEN(1,2)"
    If celF = 1 Then
        celF = "Бытовая техника"
        celF.Offset(0, -4) = WorksheetFunction.RandBetween(ThisWorkbook.Worksheets(NumWS).Range("C2") - 100, ThisWorkbook.Worksheets(NumWS).Range("D19"))
        celF.Offset(0, -2) = "HAN"
        intHAN = celF.Offset(0, -4)
        For Each celHAN In ThisWorkbook.Worksheets(NumWS).Range("$A$22:$A$35")
            cHAN = celHAN.Value
            cHN = celHAN.Offset(0, 1).Value
            If intHAN >= cHAN And intHAN <= cHN Then
                celF.Offset(0, -3) = celHAN.Offset(-1, 2)
                Exit For
            Else
                celF.Offset(0, -3) = celHAN.Offset(0, 2)
            End If
        Next celHAN
    ElseIf celF = 2 Then
        celF = "Дом и сад"
        celF.Offset(0, -4) = WorksheetFunction.RandBetween(Int(ThisWorkbook.Worksheets(NumWS).Range("A22")) - 100, Int(ThisWorkbook.Worksheets(NumWS).Range("B35")))
        celF.Offset(0, -2) = "NEWPROMO"
        intNEW = celF.Offset(0, -4)
        For Each celNEW In ThisWorkbook.Worksheets(NumWS).Range("$A$22:$A$35")
            cNEW = celNEW.Value
            cNW = celNEW.Offset(0, 1).Value
            If intNEW >= cNEW And intNEW <= cNW Then
                celF.Offset(0, -3) = celNEW.Offset(0, 2)
                Exit For
            Else
                celF.Offset(0, -3) = celNEW.Offset(0, 2)
            End If
        Next celNEW
    End If
Next celF

End Sub
