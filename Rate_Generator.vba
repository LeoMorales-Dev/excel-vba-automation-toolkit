Sub Rate_Generator()
    ' --- VERSION ANONIMIZADA PARA PORTAFOLIO ---
    ' Descripcion: Filtra y estructura tarifas masivas para carga en sistema central.
    
    Dim NuevoArchivo As Workbook: Set NuevoArchivo = Workbooks.Add
    Dim ValorAL2 As String: ValorAL2 = ThisWorkbook.Sheets("TARIFAS").Range("AL2").Value
    Dim UltimaFila As Long: UltimaFila = ThisWorkbook.Sheets("CARGA CARS").Cells(Rows.Count, "A").End(xlUp).Row
    Dim RangoValores As Range: Set RangoValores = ThisWorkbook.Sheets("CARGA CARS").Range("B2:B" & UltimaFila)
    Dim RangoFiltrado As Range, celda As Range

    ' Logica de segmentacion por Prefijo de Locacion y Marca
    For Each celda In RangoValores
        Dim Marca As String: Marca = ThisWorkbook.Sheets("CARGA CARS").Cells(celda.Row, 1).Value
        Select Case True
            Case Left(ValorAL2, 1) = "D" And Marca = "BRAND_D" ' Dollar Anom.
            Case Left(ValorAL2, 1) = "T" And Marca = "BRAND_T" ' Thrifty Anom.
            Case Left(ValorAL2, 1) = "F" And Marca = "BRAND_F" ' Firefly Anom.
            Case Len(ValorAL2) < 4 And Marca = "BRAND_H"      ' Hertz Anom.
            Case Else: GoTo Saltar
        End Select

        If RangoFiltrado Is Nothing Then Set RangoFiltrado = celda Else Set RangoFiltrado = Union(RangoFiltrado, celda)
Saltar:
    Next celda

    ' Procesamiento de datos financieros
    Dim rowDest As Long: rowDest = 2
    For Each celda In RangoFiltrado
        With NuevoArchivo.Sheets(1)
            .Cells(rowDest, 1).Value = ValorAL2
            .Cells(rowDest, 2).Value = celda.Value
            ' Redondeo financiero de precision
            For Col = 6 To 10
                If IsNumeric(celda.Offset(0, Col - 2).Value) Then
                    .Cells(rowDest, Col).Value = Round(celda.Offset(0, Col - 2).Value, 2)
                End If
            Next Col
            rowDest = rowDest + 1
        End With
    Next celda

    MsgBox "Generacion de Tarifas completada.", vbInformation
End Sub
