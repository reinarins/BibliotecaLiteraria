Sub CustomizarExcel()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rng As Range
    Dim i As Long


' Para cada worksheet, autoajusta el ancho de todas las columnas

    For Each ws In ActiveWorkbook.Worksheets
        If Not ws.ProtectContents Then 
            ws.Cells.EntireColumn.AutoFit

            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            If lastRow >= 1 And lastCol >= 1 Then
                Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
            
                ' Encabezados en negrita
                rng.Rows(1).Font.Bold = True
                rng.Rows(1).HorizontalAlignment = xlCenter
                rng.Rows(1).Interior.Color = RGB(96, 162, 221)
            
                ' Formato cebra (filas alternas)
                For i = 2 To lastRow
                    If i Mod 2 = 0 Then
                        rng.Rows(i).Interior.Color = RGB(222, 230, 253)
                    Else
                        rng.Rows(i).Interior.ColorIndex = xlNone
                    End If
                Next i

                ' Bordes
                rng.Borders.LineStyle = xlContinuous
                
            End If
        End If
    Next ws

End Sub