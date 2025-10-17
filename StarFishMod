Sub ProcessExcel()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' 1. Modify column A (Remove last 2 digits, except A1)
    Dim i As Long
    For i = 2 To lastRow
        If Len(ws.Cells(i, 1).Value) > 2 Then
            ws.Cells(i, 1).Value = Left(ws.Cells(i, 1).Value, Len(ws.Cells(i, 1).Value) - 2)
        End If
    Next i

    ' 2. Delete specified columns (F-R, T-BZ)
    ws.Range("F:R,T:BZ").Delete Shift:=xlToLeft

    ' 3. Insert & Expand Formula in G Column
    ws.Cells(2, 7).Formula = "=CONCATENATE(A2,""SP"")"
    ws.Range("G2:G" & lastRow).FillDown

    ' 4. Insert & Expand Formula in H Column
    ws.Cells(2, 8).Formula = "=G2&""-""&""MON""&""-""&C2&D2&""-""&B2"
    ws.Range("H2:H" & lastRow).FillDown

    ' 5. Copy Column H values to Column I (Keep Values Only)
    ws.Range("H2:H" & lastRow).Copy
    ws.Range("I2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' 6. Enable filtering on Column F
    ws.Range("F1:F" & lastRow).AutoFilter

    ' 7. Resize all columns
    ws.Columns.AutoFit

    MsgBox "Processing Complete!", vbInformation, "Excel Automation"
End Sub

