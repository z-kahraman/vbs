' Excel uygulamasını başlat
Set objExcel = CreateObject("Excel.Application")

' Excel dosyasını aç
Set objWorkbook = objExcel.Workbooks.Open("C:\UpWork\ExcelProfit\test.xlsx")

' Çalışma sayfasını seç
Set objWorksheet = objWorkbook.Worksheets("Sayfa1") ' Çalışma sayfasının adını değiştirin


' Verilerin bulunduğu sütunu belirleyin
Set columnB = objWorksheet.Columns("B")

' Sütundaki son dolu hücrenin satırını bulun
lastRow = objWorksheet.Cells(objWorksheet.Rows.Count, columnB.Column).End(-4162).Row

' Her hücreyi döngü ile kontrol edin ve değeri değiştirin
For i = 1 To lastRow
    If Not IsEmpty(columnB.Cells(i)) Then
        ' Değeri değiştirin: . yerine , kullanın
        columnB.Cells(i).Value = Replace(columnB.Cells(i).Value, ".", ",")
    End If
Next

Set columnC = objWorksheet.Columns("C")

' Sütundaki son dolu hücrenin satırını bulun
lastRow = objWorksheet.Cells(objWorksheet.Rows.Count, columnC.Column).End(-4162).Row

' Her hücreyi döngü ile kontrol edin ve değeri değiştirin
For i = 1 To lastRow
    If Not IsEmpty(columnC.Cells(i)) Then
        ' Değeri değiştirin: . yerine , kullanın
        columnC.Cells(i).Value = Replace(columnC.Cells(i).Value, ".", ",")
    End If
Next

' Verilerin olduğu sütunları belirle
columnB = "B" ' İlk sütunun harf değeri
columnC = "C" ' İkinci sütunun harf değeri
targetColumn = "F" ' Hedef sütunun harf değerini belirleyin
objWorksheet.Range(targetColumn & 1).Value = "Profit"

' İlk satırın numarası
startRow = 2 ' İlk satırı değiştirin (başlık satırı varsa 2, yoksa 1)

' Son satırın numarasını bul
lastRow = objWorksheet.Cells(objWorksheet.Rows.Count, columnB).End(-4162).Row ' -4162: xlUp

' Profit hesaplama
For i = startRow To lastRow

    ' Sütun A'daki hücre adresini al
    cellB = columnB & i
    
    ' Sütun B'deki hücre adresini al
    cellC = columnC & i
    
    ' Profit formülünü oluştur
    formula = "=" & cellC & "-" & cellB


    ' Formülü hedef hücreye yazdır
    targetCell = targetColumn & i
    objWorksheet.Range(targetCell).Formula = formula

     ' Sütun B ve C için Text to Columns işlemi
    objWorksheet.Range(cellB).NumberFormat = "General"
    objWorksheet.Range(cellC).NumberFormat = "General"
Next

' Excel dosyasını kaydet
objWorkbook.Save

' Excel uygulamasını kapat
objWorkbook.Close
objExcel.Quit

' Belleği temizle
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing