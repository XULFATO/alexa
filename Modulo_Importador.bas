Attribute VB_Name = "Modulo_Importador"
Option Explicit

' ============================================================
'  IMPORTADOR MAESTRO / ESCLAVO  v3
'  - Maestro : CSV separador ";"
'              campo clave: columna que contenga "NISS" (texto, ceros izq.)
'  - Esclavo : XLSX
'              Fila 1       : labels  (NIC CODE, Name, Nif...)
'              Filas 2-5    : una de ellas tiene codigos A001, A002...
'              Fila 6+      : datos
'              campo clave  : columna B (o la que tenga "NIC CODE" en fila 1)
'  - Output  : libro activo, dos hojas ("MAESTRO" y "ESCLAVO")
'              ambas hojas con headers identicos (= headers del MAESTRO)
'              esclavo reordenado segun maestro; columnas sin match descartadas
' ============================================================

Public Sub ImportarMaestroEsclavo()

    Dim sRutaMaestro As String
    Dim sRutaEsclavo As String

    sRutaMaestro = SeleccionarFichero("Selecciona el CSV Maestro", "CSV (*.csv),*.csv")
    If sRutaMaestro = "" Then MsgBox "Cancelado.", vbInformation: Exit Sub

    sRutaEsclavo = SeleccionarFichero("Selecciona el Excel Esclavo", "Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm")
    If sRutaEsclavo = "" Then MsgBox "Cancelado.", vbInformation: Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    ' ---- PASO 1: CSV Maestro ----
    Dim arrMaestro() As String
    Dim nFilasMaestro As Long
    arrMaestro = LeerCSV(sRutaMaestro, nFilasMaestro)

    Dim iHeaderMaestro As Long
    iHeaderMaestro = BuscarFilaHeader(arrMaestro, nFilasMaestro)
    If iHeaderMaestro < 0 Then
        MsgBox "No se encontro fila de headers en el Maestro.", vbCritical: GoTo Salir
    End If

    Dim arrHeaderMaestro() As String
    arrHeaderMaestro = SplitCSVLine(arrMaestro(iHeaderMaestro))
    Dim nColsMaestro As Long
    nColsMaestro = UBound(arrHeaderMaestro) + 1

    Dim iNissMaestro As Long
    iNissMaestro = BuscarColumnaEnArray(arrHeaderMaestro, "NISS")
    If iNissMaestro < 0 Then
        MsgBox "No se encontro columna 'NISS' en el Maestro.", vbCritical: GoTo Salir
    End If

    ' ---- PASO 2: Excel Esclavo ----
    Dim wbEsclavo As Workbook
    Set wbEsclavo = Workbooks.Open(sRutaEsclavo, ReadOnly:=True)
    Dim wsEsclavo As Worksheet
    Set wsEsclavo = wbEsclavo.Sheets(1)

    ' Buscar fila de labels (fila 1 segun estructura conocida, pero buscar dinamicamente)
    ' y fila de codigos A001, A002... (entre filas 1-5)
    Dim iFilaCodigos As Long, iFilaLabels As Long
    BuscarFilasEsclavo wsEsclavo, iFilaCodigos, iFilaLabels

    If iFilaCodigos < 0 Then
        wbEsclavo.Close False
        MsgBox "No se encontro fila de codigos (Axxx) en el Esclavo.", vbCritical: GoTo Salir
    End If
    If iFilaLabels < 0 Then
        wbEsclavo.Close False
        MsgBox "No se encontro fila de labels en el Esclavo.", vbCritical: GoTo Salir
    End If

    ' Buscar columna NIC CODE en fila de labels (B1 segun estructura)
    Dim nColsEsclavo As Long
    nColsEsclavo = wsEsclavo.UsedRange.Columns.Count

    Dim arrCodEsclavo() As String   ' codigos: A001, A002...
    Dim arrLblEsclavo() As String   ' labels: NIC CODE, Name, Nif...
    ReDim arrCodEsclavo(1 To nColsEsclavo)
    ReDim arrLblEsclavo(1 To nColsEsclavo)
    Dim c As Long
    For c = 1 To nColsEsclavo
        arrCodEsclavo(c) = Trim(CStr(wsEsclavo.Cells(iFilaCodigos, c).Value))
        arrLblEsclavo(c) = Trim(CStr(wsEsclavo.Cells(iFilaLabels, c).Value))
    Next c

    ' Buscar NIC CODE en labels (flexible: sin espacios)
    Dim iNikkodeEsclavo As Long
    iNikkodeEsclavo = BuscarColumnaFlexible(arrLblEsclavo, "NICCODE", nColsEsclavo)
    If iNikkodeEsclavo < 0 Then iNikkodeEsclavo = BuscarColumnaFlexible(arrCodEsclavo, "NICCODE", nColsEsclavo)
    If iNikkodeEsclavo < 0 Then iNikkodeEsclavo = BuscarColumnaEnArrayBase1(arrLblEsclavo, "NIC", nColsEsclavo)
    If iNikkodeEsclavo < 0 Then
        wbEsclavo.Close False
        MsgBox "No se encontro 'NIC CODE' en el Esclavo." & vbCrLf & _
               "Fila de labels detectada: " & iFilaLabels & vbCrLf & _
               "Primeras 5 celdas de esa fila: " & _
               wsEsclavo.Cells(iFilaLabels, 1).Value & " | " & _
               wsEsclavo.Cells(iFilaLabels, 2).Value & " | " & _
               wsEsclavo.Cells(iFilaLabels, 3).Value & " | " & _
               wsEsclavo.Cells(iFilaLabels, 4).Value & " | " & _
               wsEsclavo.Cells(iFilaLabels, 5).Value, vbCritical
        GoTo Salir
    End If

    ' ---- PASO 3: Mapa de reordenacion CA001->A001 ----
    Dim arrMapEsclavo() As Long
    ReDim arrMapEsclavo(0 To nColsMaestro - 1)
    Dim i As Long
    For i = 0 To nColsMaestro - 1
        Dim sCodMaestro As String
        sCodMaestro = Trim(arrHeaderMaestro(i))
        Dim sBuscado As String
        If UCase(Left(sCodMaestro, 1)) = "C" Then
            sBuscado = Mid(sCodMaestro, 2)   ' CA001 -> A001
        Else
            sBuscado = sCodMaestro
        End If
        arrMapEsclavo(i) = BuscarColumnaEnArrayBase1(arrCodEsclavo, sBuscado, nColsEsclavo)
    Next i

    ' ---- PASO 4: Hojas de salida ----
    Dim wsM As Worksheet, wsE As Worksheet
    Set wsM = ObtenerOCrearHoja(ThisWorkbook, "MAESTRO")
    Set wsE = ObtenerOCrearHoja(ThisWorkbook, "ESCLAVO")
    wsM.Cells.ClearContents
    wsE.Cells.ClearContents
    wsM.Cells.NumberFormat = "@"
    wsE.Cells.NumberFormat = "@"

    ' ---- PASO 5: Headers identicos en ambas hojas (headers del MAESTRO) ----
    Dim j As Long
    For j = 0 To nColsMaestro - 1
        wsM.Cells(1, j + 1).NumberFormat = "General"
        wsM.Cells(1, j + 1).Value = arrHeaderMaestro(j)
        wsE.Cells(1, j + 1).NumberFormat = "General"
        wsE.Cells(1, j + 1).Value = arrHeaderMaestro(j)
    Next j
    wsM.Rows(1).Font.Bold = True
    wsE.Rows(1).Font.Bold = True

    ' ---- PASO 6: Datos Maestro ----
    Dim iFilaSalidaM As Long
    iFilaSalidaM = 2
    Dim r As Long
    For r = iHeaderMaestro + 1 To nFilasMaestro - 1
        If Len(Trim(arrMaestro(r))) = 0 Then GoTo SiguienteFilaMaestro
        Dim arrFila() As String
        arrFila = SplitCSVLine(arrMaestro(r))
        For j = 0 To nColsMaestro - 1
            If j <= UBound(arrFila) Then
                wsM.Cells(iFilaSalidaM, j + 1).Value = arrFila(j)
            End If
        Next j
        iFilaSalidaM = iFilaSalidaM + 1
SiguienteFilaMaestro:
    Next r

    ' ---- PASO 7: Datos Esclavo reordenado ----
    ' Los datos empiezan en la fila siguiente a la mayor de iFilaCodigos/iFilaLabels
    Dim iInicioData As Long
    iInicioData = iFilaCodigos
    If iFilaLabels > iInicioData Then iInicioData = iFilaLabels
    iInicioData = iInicioData + 1

    Dim iUltimaFilaE As Long
    iUltimaFilaE = wsEsclavo.Cells(wsEsclavo.Rows.Count, iNikkodeEsclavo).End(xlUp).Row

    Dim iFilaSalidaE As Long
    iFilaSalidaE = 2
    Dim rE As Long
    Dim iColE As Long
    For rE = iInicioData To iUltimaFilaE
        For j = 0 To nColsMaestro - 1
            iColE = arrMapEsclavo(j)
            If iColE > 0 Then
                wsE.Cells(iFilaSalidaE, j + 1).Value = _
                    Trim(CStr(wsEsclavo.Cells(rE, iColE).Value))
            End If
        Next j
        iFilaSalidaE = iFilaSalidaE + 1
    Next rE

    wbEsclavo.Close False

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Importacion completada." & vbCrLf & _
           "  Maestro : " & iFilaSalidaM - 2 & " filas  ->  hoja MAESTRO" & vbCrLf & _
           "  Esclavo : " & iFilaSalidaE - 2 & " filas  ->  hoja ESCLAVO" & vbCrLf & vbCrLf & _
           "  Col. NISS     (Maestro) : col. " & iNissMaestro + 1 & vbCrLf & _
           "  Col. NIC CODE (Esclavo) : col. " & iNikkodeEsclavo & vbCrLf & _
           "  Fila labels  (Esclavo)  : fila " & iFilaLabels & vbCrLf & _
           "  Fila codigos (Esclavo)  : fila " & iFilaCodigos & vbCrLf & _
           "  Inicio datos (Esclavo)  : fila " & iInicioData, vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
Salir:
End Sub

' ============================================================
'  HELPERS
' ============================================================

Private Function SeleccionarFichero(sTitulo As String, sFiltro As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = sTitulo
        .Filters.Clear
        .Filters.Add sFiltro, Split(sFiltro, ",")(1)
        .AllowMultiSelect = False
        If .Show = -1 Then
            SeleccionarFichero = .SelectedItems(1)
        Else
            SeleccionarFichero = ""
        End If
    End With
End Function

Private Function LeerCSV(sRuta As String, ByRef nFilas As Long) As String()
    Dim iFile As Integer
    iFile = FreeFile
    Open sRuta For Input As #iFile
    Dim sLinea As String
    Dim arr() As String
    ReDim arr(0 To 9999)
    nFilas = 0
    Do While Not EOF(iFile)
        Line Input #iFile, sLinea
        If nFilas > UBound(arr) Then ReDim Preserve arr(0 To UBound(arr) + 5000)
        arr(nFilas) = sLinea
        nFilas = nFilas + 1
    Loop
    Close #iFile
    LeerCSV = arr
End Function

' Divide linea CSV por ";" respetando campos entre comillas dobles
Private Function SplitCSVLine(sLinea As String) As String()
    Dim arrResult() As String
    ReDim arrResult(0 To 199)
    Dim nCols As Long
    nCols = 0
    Dim sCampo As String
    sCampo = ""
    Dim bEnComillas As Boolean
    bEnComillas = False
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(sLinea)
        ch = Mid(sLinea, i, 1)
        If ch = """" Then
            bEnComillas = Not bEnComillas
        ElseIf ch = ";" And Not bEnComillas Then
            If nCols > UBound(arrResult) Then ReDim Preserve arrResult(0 To UBound(arrResult) + 100)
            arrResult(nCols) = sCampo
            nCols = nCols + 1
            sCampo = ""
        Else
            sCampo = sCampo & ch
        End If
    Next i
    If nCols > UBound(arrResult) Then ReDim Preserve arrResult(0 To nCols)
    arrResult(nCols) = sCampo
    ReDim Preserve arrResult(0 To nCols)
    SplitCSVLine = arrResult
End Function

Private Function BuscarFilaHeader(arr() As String, nFilas As Long) As Long
    Dim i As Long
    For i = 0 To nFilas - 1
        If Len(Trim(arr(i))) > 0 Then
            BuscarFilaHeader = i
            Exit Function
        End If
    Next i
    BuscarFilaHeader = -1
End Function

' 0-based, parcial, case-insensitive
Private Function BuscarColumnaEnArray(arr() As String, sBuscar As String) As Long
    Dim i As Long
    For i = 0 To UBound(arr)
        If InStr(1, UCase(arr(i)), UCase(sBuscar)) > 0 Then
            BuscarColumnaEnArray = i
            Exit Function
        End If
    Next i
    BuscarColumnaEnArray = -1
End Function

' 1-based, exacta primero luego parcial
Private Function BuscarColumnaEnArrayBase1(arr() As String, sBuscar As String, nCols As Long) As Long
    Dim i As Long
    For i = 1 To nCols
        If UCase(Trim(arr(i))) = UCase(Trim(sBuscar)) Then
            BuscarColumnaEnArrayBase1 = i
            Exit Function
        End If
    Next i
    For i = 1 To nCols
        If InStr(1, UCase(arr(i)), UCase(sBuscar)) > 0 Then
            BuscarColumnaEnArrayBase1 = i
            Exit Function
        End If
    Next i
    BuscarColumnaEnArrayBase1 = -1
End Function

' 1-based, elimina espacios antes de comparar (NIC CODE -> NICCODE)
Private Function BuscarColumnaFlexible(arr() As String, sBuscar As String, nCols As Long) As Long
    Dim i As Long
    For i = 1 To nCols
        If UCase(Replace(Trim(arr(i)), " ", "")) = UCase(sBuscar) Then
            BuscarColumnaFlexible = i
            Exit Function
        End If
    Next i
    BuscarColumnaFlexible = -1
End Function

Private Sub BuscarFilasEsclavo(ws As Worksheet, ByRef iFilaCodigos As Long, ByRef iFilaLabels As Long)
    ' Busca entre las primeras 10 filas:
    '   iFilaCodigos : fila con patron A+digitos (A001, A002...)
    '   iFilaLabels  : fila anterior a iFilaCodigos (labels descriptivos)
    iFilaCodigos = -1
    iFilaLabels = -1
    Dim nCols As Long
    nCols = ws.UsedRange.Columns.Count
    Dim r As Long, c As Long
    For r = 1 To 10
        Dim nMatch As Long
        nMatch = 0
        For c = 1 To nCols
            Dim sVal As String
            sVal = Trim(CStr(ws.Cells(r, c).Value))
            If Len(sVal) >= 2 Then
                If UCase(Left(sVal, 1)) = "A" And IsNumeric(Mid(sVal, 2)) Then
                    nMatch = nMatch + 1
                End If
            End If
        Next c
        If nMatch >= 3 Then
            iFilaCodigos = r
            ' Labels siempre en la fila anterior
            If r > 1 Then
                iFilaLabels = r - 1
            Else
                iFilaLabels = -1  ' no hay fila anterior, se usara la siguiente
            End If
            Exit For
        End If
    Next r
End Sub

Private Function ObtenerOCrearHoja(wb As Workbook, sNombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sNombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sNombre
    End If
    Set ObtenerOCrearHoja = ws
End Function
