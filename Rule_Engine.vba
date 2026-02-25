Sub Rule_Engine()
    ' --- VERSION ANONIMIZADA PARA PORTAFOLIO ---
    ' Descripcion: Valida campos y asigna reglas de negocio segun el tipo de tarifa.
    
    Dim NuevoArchivo As Workbook
    Dim RangoDatos As Range, celda As Range
    Dim DiccionarioUnicos As Object
    Dim FilaDestino As Long
    Dim ValorLoc As Variant, ValorEffDate As Variant, ValorEffTime As Variant
    Dim Codigo As String, MensajeError As String

    ' Validacion de integridad de datos
    MensajeError = ""
    If ThisWorkbook.Sheets("TARIFAS").Range("AL2").Value = "" Then MensajeError = MensajeError & "Locacion" & vbCrLf
    If ThisWorkbook.Sheets("TARIFAS").Range("AL5").Value = "" Then MensajeError = MensajeError & "Rate_eff_date" & vbCrLf
    If ThisWorkbook.Sheets("TARIFAS").Range("AL8").Value = "" Then MensajeError = MensajeError & "Rate_eff_time" & vbCrLf

    If MensajeError <> "" Then
        MsgBox "Campos incompletos:" & vbCrLf & MensajeError, vbCritical, "Error de Validacion"
        Exit Sub
    End If

    ' Captura de parametros globales
    With ThisWorkbook.Sheets("TARIFAS")
        ValorLoc = .Range("AL2").Value
        ValorEffDate = .Range("AL5").Value
        ValorEffTime = .Range("AL8").Value
    End With

    Set NuevoArchivo = Workbooks.Add
    ' Copiar estructura de encabezados
    ThisWorkbook.Sheets("Macro").Range("B5:DK5").Copy
    NuevoArchivo.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteValues

    Set DiccionarioUnicos = CreateObject("Scripting.Dictionary")
    With ThisWorkbook.Sheets("CARGA CARS")
        Set RangoDatos = .Range("B2:B" & .Cells(.Rows.Count, "B").End(xlUp).Row)
    End With

    FilaDestino = 2
    For Each celda In RangoDatos
        If celda.Value <> "" And Not DiccionarioUnicos.exists(CStr(celda.Value)) Then
            DiccionarioUnicos.Add CStr(celda.Value), True
            With NuevoArchivo.Sheets(1)
                Codigo = UCase(celda.Value)
                .Cells(FilaDestino, 1).Value = ValorLoc
                .Cells(FilaDestino, 2).Value = celda.Value
                
                ' Lógica de Negocio Anonimizada (Mapeo de Tarifas)
                Select Case Codigo
                    Case "AVAD":  .Cells(FilaDestino, 36).Value = "RETAIL_PARTNER_TYPE_A_MXN"
                    Case "MPRD":  .Cells(FilaDestino, 36).Value = "PROMO_DISCOUNT_USD"
                    Case "AFLX":  .Cells(FilaDestino, 36).Value = "BROKER_INCLUSIVE_RATE_LDW"
                    Case "AM375": .Cells(FilaDestino, 36).Value = "AIRLINE_PARTNER_SPECIAL_RATE"
                    Case "CITI":  .Cells(FilaDestino, 36).Value = "PREMIUM_BANKING_LOYALTY_RATE"
                    Case "MTRVL": .Cells(FilaDestino, 36).Value = "TRAVEL_AGENCY_OFFER_20"
                    Case Else:    .Cells(FilaDestino, 36).Value = "STANDARD_RETAIL_RATE"
                End Select
                
                ' Configuracion de constantes de sistema
                .Cells(FilaDestino, 22).Value = "Y": .Cells(FilaDestino, 24).Value = "Y"
                .Cells(FilaDestino, 28).Value = 6
            End With
            FilaDestino = FilaDestino + 1
        End If
    Next celda

    ' Exportacion final a CSV
    Dim RutaDestino As Variant
    RutaDestino = Application.GetSaveAsFilename(InitialFileName:="PROCESSED_RULES.csv", FileFilter:="CSV Files (*.csv), *.csv")
    If RutaDestino <> False Then
        NuevoArchivo.SaveAs Filename:=RutaDestino, FileFormat:=xlCSV, Local:=True
        MsgBox "Proceso completado con exito.", vbInformation
    End If
End Sub
