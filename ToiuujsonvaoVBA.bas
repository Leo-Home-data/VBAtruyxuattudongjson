Attribute VB_Name = "ToiuujsonvaoVBA"
Sub ProcessDataOptimized()
    Dim ws As Worksheet
    Dim rngData As Range
    Dim lastRow As Long
    Dim data As Variant
    Dim total As Double
    Dim i As Long

    ' Thiet lap worksheet và xác dinh vung du lieu
    Set ws = ThisWorkbook.Sheets("Khach hang")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).row
    Set rngData = ws.Range("D1:H" & lastRow)

    'Doc du lieu vao mang
    data = rngData.Value

    ' Tinh tong so tien cho các khách hàng có "Loai dich v?" là "rut tien"
    total = 0
    For i = 2 To UBound(data, 1) ' Bat dau tu dong 2 bo qua tieu de
        If data(i, 4) = "rut tien" And IsNumeric(data(i, 2)) Then
            total = total + data(i, 2)
        End If
    Next i

    ' Hien thi ket qua
    MsgBox "Tong so tien rut duoc la: " & total

    ' Bat tinh nang toi uu hoa
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


