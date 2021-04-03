Attribute VB_Name = "Module1"
Sub stonks():

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

Dim ticker As String

Dim yearly_change As Double
Dim summary_row As Double
Dim volume_total As Double
Dim volume As Double


Dim last_stonk As Double
Dim first_stonk As Double


first_stonk = Cells(2, 3).Value


summary_row = 2
volume_total = 0


For i = 2 To RowCount


volume_total = volume_total + Cells(i, 7).Value

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        
        last_stonk = Cells(i, 6).Value
        
        Cells(summary_row, "I").Value = ticker
        Cells(summary_row, "J").Value = volume_total
        Cells(summary_row, "K").Value = last_stonk - first_stonk
            If first_stonk = 0 Then
                Cells(summary_row, "L").Value = "yolo infinity lmao"
            Else
                Cells(summary_row, "L").Value = ((last_stonk - first_stonk) / (first_stonk))
            End If
        first_stonk = Cells(i + 1, 3).Value


        
        summary_row = summary_row + 1
        volume_total = 0
        
        
    End If
Next i

MsgBox (RowCount)

For i = 2 To RowCount


    If Cells(i, "K").Value >= 0 Then
        Cells(i, "K").Interior.Color = vbGreen
    Else
        Cells(i, "K").Interior.Color = vbRed
    End If
    
Cells(i, "L").NumberFormat = "0.00%"

Next i



End Sub

