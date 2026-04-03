Attribute VB_Name = "Importador_PageV"
Option Explicit

' ============================================================
'  PROYECTO A - IMPORTADOR CSV + EXCEL ESCLAVO
'  Modulo independiente. No toca nada del proyecto B.
'
'  ImportarCSV   -> Page 1 v1  (guarda nombre en J1 del MENU)
'  ImportarExcel -> Page 1 v2  (guarda nombre en J2 del MENU)
'  ActualizarBotones -> cambia texto/color de Btn1 y Btn2
'
'  Estructura esclavo:
'    Fila 1 : labels PT  (NIC Code, Nome...)  <- cabeceras salida
'    Fila 2 : labels EN  (Employee ID, Name...)
'    Fila 3 : codigos    (A002, A001, AL11...)
'    Fila 4 : tipos dato
'    Fila 5 : descripciones
'    Fila 6+: datos
'
'  Mapeo:
'    NISS maestro -> NIC Code esclavo (col B)
'    CAxxxx       -> quitar C -> Axxxx -> buscar fila 3
'    Col 1 salida -> siempre "EMPLOYEE ID"
'    B357 / B001  -> numerico 2 decimales
' ============================================================

' ============================================================
'  IMPORTAR CSV MAESTRO -> Page 1 v1
' ============================================================

Public Sub ImportarCSV()

    Dim sRuta As String
    sRuta = SeleccionarFicheroPageV("Selecciona el CSV Maestro", "CSV (*.csv),*.csv")
    If sRuta = "" Then MsgBox "Cancelado.", vbInformation: Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    Dim arrLineas() As String
    Dim nFilas As Long
    arrLineas = LeerCSV(sRuta, nFilas)

    Dim iHdr As Long
    iHdr = BuscarFilaHeaderPageV(arrLineas, nFilas)
    If iHdr < 0 Then MsgBox "No se encontro header en el CSV.", vbCritical: GoTo Salir

    arrLineas(iHdr) = LimpiarBOM(arrLineas(iHdr))
    Dim arrHdr() As String
    arrHdr = SplitCSVLine(arrLineas(iHdr))
    Dim nCols As Long
    nCols = UBound(arrHdr) + 1
    Dim h As Long
    For h = 0 To nCols - 1
        arrHdr(h) = LimpiarBOM(Trim(arrHdr(h)))
    Next h

    Dim ws As Worksheet
    Set ws = ObtenerOCrearHojaPageV(ThisWorkbook, "Page 1 v1")
    ws.Cells.ClearContents
    ws.Cells.NumberFormat = "@"

    Dim j As Long
    For j = 0 To nCols - 1
        ws.Cells(1, j + 1).NumberFormat = "General"
        If j = 0 Then
            ws.Cells(1, j + 1).Value = "EMPLOYEE ID"
        Else
            ws.Cells(1, j + 1).Value = arrHdr(j)
        End If
        Dim sCod As String
        sCod = arrHdr(j)
        If UCase(Left(sCod, 1)) = "C" Then sCod = Mid(sCod, 2)
        If EsColumnaDecimal(sCod) Then
            ws.Columns(j + 1).NumberFormat = "0.00"
            ws.Cells(1, j + 1).NumberFormat = "General"
        End If
    Next j
    ws.Rows(1).Font.Bold = True

    Dim iSalida As Long
    iSalida = 2
    Dim r As Long
    For r = iHdr + 1 To nFilas - 1
        If Len(Trim(arrLineas(r))) = 0 Then GoTo SigFila
        Dim arrFila() As String
        arrFila = SplitCSVLine(arrLineas(r))
        For j = 0 To nCols - 1
            If j <= UBound(arrFila) Then
                Dim sCodD As String
                sCodD = arrHdr(j)
                If UCase(Left(sCodD, 1)) = "C" Then sCodD = Mid(sCodD, 2)
                ' DEBUG: descomentar para ver que codigo llega
                ' MsgBox "Col " & j & " codigo=[" & sCodD & "] decimal=" & EsColumnaDecimal(sCodD)
                If EsColumnaDecimal(sCodD) Then
                    Dim sNum As String
                    sNum = Replace(arrFila(j), ",", ".")
                    If IsNumeric(sNum) Then
                        ws.Cells(iSalida, j + 1).Value = CDbl(sNum)
                    Else
                        ws.Cells(iSalida, j + 1).Value = arrFila(j)
                    End If
                Else
                    ' Forzar texto para preservar ceros iniciales
                    ws.Cells(iSalida, j + 1).NumberFormat = "@"
                    ws.Cells(iSalida, j + 1).Value = arrFila(j)
                End If
            End If
        Next j
        iSalida = iSalida + 1
SigFila:
    Next r

    ' Guardar nombre en J1 del MENU para CompararHojas
    Dim wsMenuCSV As Worksheet
    Set wsMenuCSV = ThisWorkbook.Worksheets("MENU")
    wsMenuCSV.Unprotect Password:="ADP"
    wsMenuCSV.Range("J1").Value = "Page 1 v1"
    wsMenuCSV.Protect Password:="ADP", DrawingObjects:=False, Contents:=True, Scenarios:=True

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "CSV importado: " & iSalida - 2 & " filas -> 'Page 1 v1'", vbInformation
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
Salir:
End Sub

' ============================================================
'  IMPORTAR EXCEL ESCLAVO -> Page 1 v2
' ============================================================

Public Sub ImportarExcel()

    Dim wsM As Worksheet
    On Error Resume Next
    Set wsM = ThisWorkbook.Sheets("Page 1 v1")
    On Error GoTo 0
    If wsM Is Nothing Then
        MsgBox "Primero importa el CSV (Page 1 v1 no existe).", vbExclamation
        Exit Sub
    End If

    Dim sRuta As String
    sRuta = SeleccionarFicheroPageV("Selecciona el Excel Esclavo", "Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm")
    If sRuta = "" Then MsgBox "Cancelado.", vbInformation: Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    ' Leer headers de Page 1 v1
    Dim nColsM As Long
    nColsM = wsM.Cells(1, wsM.Columns.Count).End(xlToLeft).Column
    Dim arrHdrM() As String
    ReDim arrHdrM(0 To nColsM - 1)
    Dim h As Long
    For h = 0 To nColsM - 1
        arrHdrM(h) = Trim(CStr(wsM.Cells(1, h + 1).Value))
    Next h

    ' Abrir esclavo, buscar hoja Data
    Dim wbE As Workbook
    Set wbE = Workbooks.Open(sRuta, ReadOnly:=True)
    Dim wsD As Worksheet
    Dim wsTemp As Worksheet
    For Each wsTemp In wbE.Sheets
        If UCase(Trim(wsTemp.Name)) = "DATA" Then
            Set wsD = wsTemp
            Exit For
        End If
    Next wsTemp
    If wsD Is Nothing Then Set wsD = wbE.Sheets(1)

    Dim iFilaCod As Long
    iFilaCod = BuscarFilaCodigosPageV(wsD)
    If iFilaCod < 0 Then
        wbE.Close False
        MsgBox "No se encontro fila de codigos (Axxx) en el Esclavo.", vbCritical: GoTo Salir
    End If

    Dim nColsE As Long
    nColsE = wsD.UsedRange.Columns.Count
    Dim arrCodE() As String
    Dim arrLblPT() As String
    ReDim arrCodE(1 To nColsE)
    ReDim arrLblPT(1 To nColsE)
    Dim c As Long
    For c = 1 To nColsE
        arrCodE(c) = Trim(CStr(wsD.Cells(iFilaCod, c).Value))
        arrLblPT(c) = Trim(CStr(wsD.Cells(1, c).Value))
    Next c

    ' Columna NIC Code (col B)
    Dim iNicCol As Long
    iNicCol = BuscarColumnaFlexiblePageV(arrLblPT, "NICCODE", nColsE)
    If iNicCol < 0 Then iNicCol = BuscarColumnaBase1PageV(arrLblPT, "NIC", nColsE)
    If iNicCol < 0 Then iNicCol = 2

    ' Mapa de columnas
    Dim arrMap() As Long
    ReDim arrMap(0 To nColsM - 1)
    Dim i As Long
    For i = 0 To nColsM - 1
        If i = 0 Then
            arrMap(i) = iNicCol
        Else
            Dim sHdrM As String
            sHdrM = arrHdrM(i)
            Dim iFound As Long
            iFound = BuscarColumnaFlexiblePageV(arrLblPT, Replace(sHdrM, " ", ""), nColsE)
            If iFound < 0 Then iFound = BuscarColumnaBase1PageV(arrLblPT, sHdrM, nColsE)
            If iFound < 0 Then
                Dim sBusc As String
                sBusc = sHdrM
                If UCase(Left(sBusc, 1)) = "C" Then sBusc = Mid(sBusc, 2)
                iFound = BuscarColumnaBase1PageV(arrCodE, sBusc, nColsE)
            End If
            arrMap(i) = iFound
        End If
    Next i

    ' Hoja Page 1 v2
    Dim wsE As Worksheet
    Set wsE = ObtenerOCrearHojaPageV(ThisWorkbook, "Page 1 v2")
    wsE.Cells.ClearContents
    wsE.Cells.NumberFormat = "@"

    Dim j As Long
    For j = 0 To nColsM - 1
        wsE.Cells(1, j + 1).NumberFormat = "General"
        wsE.Cells(1, j + 1).Value = arrHdrM(j)
        Dim iCE As Long
        iCE = arrMap(j)
        If iCE > 0 Then
            If EsColumnaDecimal(arrCodE(iCE)) Then
                wsE.Columns(j + 1).NumberFormat = "0.00"
                wsE.Cells(1, j + 1).NumberFormat = "General"
            End If
        End If
    Next j
    wsE.Rows(1).Font.Bold = True

    ' Datos desde fila 6
    Dim iUltima As Long
    iUltima = wsD.Cells(wsD.Rows.Count, iNicCol).End(xlUp).Row
    Dim iSal As Long
    iSal = 2
    Dim rE As Long
    For rE = 6 To iUltima
        If Len(Trim(CStr(wsD.Cells(rE, iNicCol).Value))) = 0 Then GoTo SigFilaE
        For j = 0 To nColsM - 1
            iCE = arrMap(j)
            If iCE > 0 Then
                If EsColumnaDecimal(arrCodE(iCE)) Then
                    Dim sNumE As String
                    sNumE = Replace(Trim(CStr(wsD.Cells(rE, iCE).Value)), ",", ".")
                    If IsNumeric(sNumE) Then
                        wsE.Cells(iSal, j + 1).Value = CDbl(sNumE)
                    Else
                        wsE.Cells(iSal, j + 1).Value = Trim(CStr(wsD.Cells(rE, iCE).Value))
                    End If
                Else
                    ' Forzar texto para preservar ceros iniciales
                    wsE.Cells(iSal, j + 1).NumberFormat = "@"
                    wsE.Cells(iSal, j + 1).Value = Trim(CStr(wsD.Cells(rE, iCE).Value))
                End If
            End If
        Next j
        iSal = iSal + 1
SigFilaE:
    Next rE

    wbE.Close False

    ' Guardar nombre en J2 del MENU para CompararHojas
    Dim wsMenuXLS As Worksheet
    Set wsMenuXLS = ThisWorkbook.Worksheets("MENU")
    wsMenuXLS.Unprotect Password:="ADP"
    wsMenuXLS.Range("J2").Value = "Page 1 v2"
    wsMenuXLS.Protect Password:="ADP", DrawingObjects:=False, Contents:=True, Scenarios:=True

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Excel importado: " & iSal - 2 & " filas -> 'Page 1 v2'" & vbCrLf & _
           "Col. NIC Code: " & iNicCol & "  |  Fila codigos: " & iFilaCod, vbInformation
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
Salir:
End Sub

' ============================================================
'  HELPERS PRIVADOS (con sufijo PageV para evitar conflictos)
' ============================================================

Private Function EsColumnaDecimal(sCodigo As String) As Boolean
    Select Case UCase(Trim(sCodigo))
        Case "B357", "B001"
            EsColumnaDecimal = True
        Case Else
            EsColumnaDecimal = False
    End Select
End Function

Private Function LimpiarBOM(s As String) As String
    Dim sBOM As String
    sBOM = Chr(239) & Chr(187) & Chr(191)
    If Left(s, 3) = sBOM Then s = Mid(s, 4)
    Do While Left(s, 1) = Chr(255) Or Left(s, 1) = Chr(254) Or _
             Left(s, 1) = Chr(239) Or Left(s, 1) = Chr(187) Or Left(s, 1) = Chr(191)
        s = Mid(s, 2)
    Loop
    LimpiarBOM = s
End Function

Private Function SeleccionarFicheroPageV(sTitulo As String, sFiltro As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = sTitulo
        .Filters.Clear
        .Filters.Add sFiltro, Split(sFiltro, ",")(1)
        .AllowMultiSelect = False
        If .Show = -1 Then
            SeleccionarFicheroPageV = .SelectedItems(1)
        Else
            SeleccionarFicheroPageV = ""
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

Private Function SplitCSVLine(sLinea As String) As String()
    Dim arrResult() As String
    ReDim arrResult(0 To 199)
    Dim nCols As Long
    nCols = 0
    Dim sCampo As String
    sCampo = ""
    Dim bEnComillas As Boolean
    bEnComillas = False
    Dim i As Long, ch As String
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

Private Function BuscarFilaHeaderPageV(arr() As String, nFilas As Long) As Long
    Dim i As Long
    For i = 0 To nFilas - 1
        If Len(Trim(arr(i))) > 0 Then
            BuscarFilaHeaderPageV = i
            Exit Function
        End If
    Next i
    BuscarFilaHeaderPageV = -1
End Function

Private Function BuscarFilaCodigosPageV(ws As Worksheet) As Long
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
            BuscarFilaCodigosPageV = r
            Exit Function
        End If
    Next r
    BuscarFilaCodigosPageV = -1
End Function

Private Function BuscarColumnaBase1PageV(arr() As String, sBuscar As String, nCols As Long) As Long
    Dim i As Long
    For i = 1 To nCols
        If UCase(Trim(arr(i))) = UCase(Trim(sBuscar)) Then
            BuscarColumnaBase1PageV = i
            Exit Function
        End If
    Next i
    For i = 1 To nCols
        If InStr(1, UCase(arr(i)), UCase(sBuscar)) > 0 Then
            BuscarColumnaBase1PageV = i
            Exit Function
        End If
    Next i
    BuscarColumnaBase1PageV = -1
End Function

Private Function BuscarColumnaFlexiblePageV(arr() As String, sBuscar As String, nCols As Long) As Long
    Dim i As Long
    For i = 1 To nCols
        If UCase(Replace(Trim(arr(i)), " ", "")) = UCase(sBuscar) Then
            BuscarColumnaFlexiblePageV = i
            Exit Function
        End If
    Next i
    BuscarColumnaFlexiblePageV = -1
End Function

Private Function ObtenerOCrearHojaPageV(wb As Workbook, sNombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sNombre)
    On Error GoTo 0
    If ws Is Nothing Then
        ' Desproteger libro para poder añadir hoja
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sNombre
    End If
    Set ObtenerOCrearHojaPageV = ws
End Function
