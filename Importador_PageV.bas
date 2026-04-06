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
    ' Importa desde Excel (primera hoja) en vez de CSV
    ' Busca columna NISS y la renombra EMPLOYEE ID en la salida

    Dim sRuta As String
    sRuta = SeleccionarFicheroPageV("Selecciona el Excel Maestro", "Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm")
    If sRuta = "" Then MsgBox "Cancelado.", vbInformation: Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    ' Abrir Excel origen, leer primera hoja
    Dim wbOrigen As Workbook
    Set wbOrigen = Workbooks.Open(sRuta, ReadOnly:=True)
    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Sheets(1)

    ' Buscar fila de headers (primera fila no vacia)
    Dim iHdr As Long
    iHdr = 1
    Do While iHdr <= 10
        If Len(Trim(CStr(wsOrigen.Cells(iHdr, 1).Value))) > 0 Then Exit Do
        iHdr = iHdr + 1
    Loop

    ' Leer headers
    Dim nCols As Long
    nCols = wsOrigen.Cells(iHdr, wsOrigen.Columns.Count).End(xlToLeft).Column
    Dim arrHdr() As String
    ReDim arrHdr(0 To nCols - 1)
    Dim h As Long
    For h = 0 To nCols - 1
        arrHdr(h) = Trim(CStr(wsOrigen.Cells(iHdr, h + 1).Value))
    Next h

    ' Buscar columna NISS
    Dim iNiss As Long
    iNiss = -1
    Dim hh As Long
    For hh = 0 To nCols - 1
        If UCase(Trim(arrHdr(hh))) = "NISS" Then
            iNiss = hh: Exit For
        End If
    Next hh
    If iNiss < 0 Then
        For hh = 0 To nCols - 1
            If InStr(UCase(arrHdr(hh)), "NISS") > 0 Then
                iNiss = hh: Exit For
            End If
        Next hh
    End If
    If iNiss < 0 Then
        wbOrigen.Close False
        MsgBox "No se encontro columna NISS en el Excel origen.", vbCritical
        GoTo Salir
    End If

    ' Preparar hoja Page 1 v1
    Dim ws As Worksheet
    Set ws = ObtenerOCrearHojaPageV(ThisWorkbook, "Page 1 v1")
    ws.Cells.ClearContents
    ws.Cells.NumberFormat = "@"

    ' Escribir headers
    Dim j As Long
    For j = 0 To nCols - 1
        ws.Cells(1, j + 1).NumberFormat = "General"
        If j = iNiss Then
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

    ' Volcar datos
    Dim iUltima As Long
    iUltima = wsOrigen.Cells(wsOrigen.Rows.Count, iNiss + 1).End(xlUp).Row

    Dim iSalida As Long
    iSalida = 2
    Dim r As Long
    For r = iHdr + 1 To iUltima
        If Len(Trim(CStr(wsOrigen.Cells(r, iNiss + 1).Value))) = 0 Then GoTo SigFila
        For j = 0 To nCols - 1
            Dim sCodD As String
            sCodD = arrHdr(j)
            If UCase(Left(sCodD, 1)) = "C" Then sCodD = Mid(sCodD, 2)
            If EsColumnaDecimal(sCodD) Then
                Dim dNum As Double
                If IsNumeric(wsOrigen.Cells(r, j + 1).Value) Then
                    dNum = wsOrigen.Cells(r, j + 1).Value
                Else
                    dNum = ConvertirDecimalCSV(Trim(CStr(wsOrigen.Cells(r, j + 1).Value)))
                End If
                ws.Cells(iSalida, j + 1).NumberFormat = "0.00"
                ws.Cells(iSalida, j + 1).Value = dNum
            Else
                Dim sVal As String
                sVal = Trim(CStr(wsOrigen.Cells(r, j + 1).Value))
                If IsNumeric(wsOrigen.Cells(r, j + 1).Value) And Left(sVal, 1) <> "0" Then
                    ws.Cells(iSalida, j + 1).NumberFormat = "General"
                    ws.Cells(iSalida, j + 1).Value = wsOrigen.Cells(r, j + 1).Value
                Else
                    ws.Cells(iSalida, j + 1).NumberFormat = "@"
                    ws.Cells(iSalida, j + 1).Value = sVal
                End If
            End If
        Next j
        iSalida = iSalida + 1
SigFila:
    Next r

    wbOrigen.Close False

    Dim wsMenu1 As Worksheet
    Set wsMenu1 = ThisWorkbook.Worksheets("MENU")
    wsMenu1.Unprotect Password:="ADP"
    wsMenu1.Range("J1").Value = "Page 1 v1"
    wsMenu1.Protect Password:="ADP", DrawingObjects:=False, Contents:=True, Scenarios:=True

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Excel importado: " & iSalida - 2 & " filas -> 'Page 1 v1'" & vbCrLf & _
           "Columna NISS encontrada en col. " & iNiss + 1 & " -> renombrada EMPLOYEE ID", vbInformation
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
                    ' Leer .Value directamente: Excel ya lo tiene como Double correcto
                    Dim dValE As Double
                    If IsNumeric(wsD.Cells(rE, iCE).Value) Then
                        dValE = wsD.Cells(rE, iCE).Value  ' asignacion directa, sin CDbl ni CStr
                    Else
                        dValE = ConvertirDecimalCSV(Trim(CStr(wsD.Cells(rE, iCE).Value)))
                    End If
                    wsE.Cells(iSal, j + 1).NumberFormat = "0.00"
                    wsE.Cells(iSal, j + 1).Value = dValE
                Else
                    Dim sValXLS As String
                    sValXLS = Trim(CStr(wsD.Cells(rE, iCE).Value))
                    If IsNumeric(wsD.Cells(rE, iCE).Value) And Left(sValXLS, 1) <> "0" Then
                        ' Numerico sin ceros iniciales: escribir como numero
                        wsE.Cells(iSal, j + 1).NumberFormat = "General"
                        wsE.Cells(iSal, j + 1).Value = CDbl(wsD.Cells(rE, iCE).Value)
                    Else
                        ' Texto o tiene ceros iniciales: preservar como texto
                        wsE.Cells(iSal, j + 1).NumberFormat = "@"
                        wsE.Cells(iSal, j + 1).Value = sValXLS
                    End If
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

Private Function ConvertirDecimalCSV(sVal As String) As Double
    ' Convierte string con coma decimal (formato europeo CSV) a Double
    ' sin depender de la configuracion regional de VBA/Windows
    ' Ejemplo: "2024,10" -> 2024.1  /  "0,60" -> 0.6
    If Len(sVal) = 0 Then ConvertirDecimalCSV = 0: Exit Function
    Dim posComa As Integer
    posComa = InStr(sVal, ",")
    If posComa > 0 Then
        ' Tiene coma: separar parte entera y decimal
        Dim sEntero As String, sDecimal As String
        sEntero  = Left(sVal, posComa - 1)
        sDecimal = Mid(sVal, posComa + 1)
        If Len(sEntero) = 0 Then sEntero = "0"
        If Not IsNumeric(sEntero) Or Not IsNumeric(sDecimal) Then
            ConvertirDecimalCSV = 0: Exit Function
        End If
        ConvertirDecimalCSV = CDbl(CLng(sEntero)) + CDbl(CLng(sDecimal)) / (10 ^ Len(sDecimal))
    ElseIf IsNumeric(sVal) Then
        ' Sin coma: entero o punto decimal anglosaxon
        ConvertirDecimalCSV = CDbl(Val(sVal))
    Else
        ConvertirDecimalCSV = 0
    End If
End Function

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
