Attribute VB_Name = "DebugPrint"
Sub ProcessDataWithDebug()
    Dim ws As Worksheet
    Dim rngData As Range
    Dim lastRow As Long
    Dim total As Double
    Dim i As Long

    ' Thiet lap worksheet và xac dinh vung du lieu
    Set ws = ThisWorkbook.Sheets("Khach hang")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).row
    Set rngData = ws.Range("D1:H" & lastRow)

    ' Tinh tong tien cua cac khach hang co loai dich vu là "rut tien"
    total = 0
    For i = 2 To lastRow ' Bat dau tu dong 2 bo qua tieu de
        If ws.Cells(i, 7).Value = "rut tien" And IsNumeric(ws.Cells(i, 5).Value) Then
            total = total + ws.Cells(i, 5).Value
        End If
        Debug.Print "Dòng " & i & ": Tong = " & total ' In gia tri cua total trong Immediate Window
    Next i

    ' Hien ket qua
    MsgBox "Tong so tien  là: " & total
End Sub


