Sub CBX_System_Processor()
    ' --- VERSION ANONIMIZADA PARA PORTAFOLIO ---
    ' Descripcion: Procesa y fragmenta datos para la plataforma Crossborder Xpress.

    Dim wsSource As Worksheet: Set wsSource = ThisWorkbook.Sheets("CARGA TSD CBX")
    
    ' Validacion de integridad: Previene la carga de valores negativos o nulos
    Dim iRow As Long
    For iRow = 2 To 126
        Dim valorI As Variant: valorI = wsSource.Cells(iRow, 9).Value
        If IsEmpty(valorI) Or valorI = "" Or (IsNumeric(valorI) And valorI < 0) Then
            MsgBox "Error de Integridad: Valores detectados en Columna I. Revise el llenado.", vbExclamation
            Exit Sub
        End If
    Next iRow

    ' Creacion de archivos de salida para integracion
    Dim wbNew As Workbook: Set wbNew = Workbooks.Add
    ' Aqui se implementa la logica de separacion por tipo (CBX, DCBX, TCBX)
    ' ... (Estructura de guardado anonimizada usando wbSource.Path)
    
    MsgBox "Archivos de integracion generados correctamente en la carpeta del proyecto.", vbInformation
End Sub
