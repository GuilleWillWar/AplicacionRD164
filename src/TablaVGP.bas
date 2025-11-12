Attribute VB_Name = "TablaVGP"
Option Explicit

' ============================================================
'   modVGP_Resumen_Directo (versiÃ³n bloques / OpciÃ³n B)
'   - Duplica plantilla VGP y rellena "Tabla resumen"
'   - Mapeo directo Interfaz -> celdas destino
'   - Soporta:
'       a) Hoja _MAP_VGP (Campo | Origen | Destino)
'       b) Mapeo interno (GetStaticMappings) si no existe la hoja
'   - Evita el lÃ­mite de 24 continuaciones con bloques + ConcatArrays
'   - GetValueByRef admite expresiones simples (p.ej. "F24/2" o "F30/F31")
' ============================================================

' ===========
'  CONFIG
' ===========
Private Const TEMPLATE_SHEET_NAME As String = "Tabla Vacia"
Private Const LOCAL_TEMPLATE_SHEET As String = "TEMPLATE_TABLA_RESUMEN"
Private Const OUTPUT_SHEET_NAME As String = "Tabla resumen"
Private Const INTERFAZ_SHEET As String = "Interfaz"
Private Const MAP_SHEET As String = "_MAP_VGP"

' ===========
'  CONSTANTES DE ORIGEN (Interfaz)
'   OJO: estas deben ser direcciones vÃ¡lidas en Interfaz,
'   salvo que quieras usar expresiones (p. ej. "F24/2")
' ===========
Private Const celdaSuperficieSector As String = "F3"
Private Const celdaQs As String = "D2"
Private Const celdaActividad As String = "D3"
Private Const celdaTipo As String = "B2"
Private Const celdaNri As String = "F2"
Private Const celdaFactor As String = "F3"
Private Const celdaMaxSuperficieAdmisible As String = "F8"
Private Const celdaRasante As String = "B3"
Private Const celdaFlag50FachadaAccesible As String = "B9"
Private Const celdaFlagRociadores As String = "B10"
Private Const celdaFlagSinLimiteTipoCSistemaFijo As String = "B12"
Private Const celdaSuperficieMaximaCorregida As String = "F8"
Private Const celdaFlagBaterias As String = "B16"
Private Const celdaFlagEstructurasIndependientes As String = "B18"
Private Const celdaFlag1Bie As String = "B21"
Private Const celdaElementosDelimitadores As String = "F9"
Private Const celdaElementosDelimiradoresCorregida As String = "F9" ' (sic)
Private Const celdaPuertasPasoPeatoneal As String = "F10"            ' (sic)
Private Const celdaParedesYTechos As String = "F14"
Private Const celdaSuelos As String = "F15"
Private Const celdaElementosSeparadores As String = "F24"
Private Const celdaElementosSeparadoresColindantes As String = "F24"
Private Const celdaEncuentrosDeFachadasYCubiertasSectores As String = "F24"  ' expresiÃ³n admitida
Private Const celdaEncuentrosDeFachadasYCubiertasEstablecimientos As String = "F24/2" ' expresiÃ³n
Private Const celdaReaccionAlFuegoDeFachadas As String = "F25"
Private Const celdaLongitudRecorridoSalidaExistente As String = "D28" ' expresiÃ³n
Private Const celdaLongitudRecorridoUnaSalidaOSinAlternativa As String = "F30" ' expresiÃ³n
Private Const celdaLongitudRecorridoCorregidaUnaSalidaOSinAlternativa As String = "F31"
Private Const celdaLongitudRecorrido2Salidas As String = "F32"
Private Const celdaFlagSistemaFijo As String = "B31"
Private Const celdaFlagStech As String = "B30"
Private Const celdaFactor2SalidasSistemaFijoHnave8 As String = "B32"
Private Const celdaLongitudRecorridoCorregidaDosSalidas As String = "F32"
Private Const celdaEstructuraPortante As String = "F46"
Private Const celdaFlagCubiertaLigeraIndependiente As String = "B47"
Private Const celdaFlagCubiertaLigeraConSistemaFijo As String = "B46"
Private Const celdaDeteccionAutomatica As String = "A51"
Private Const celdaDeteccionManual As String = "B51"
Private Const celdaExtintores As String = "C53"
Private Const celdaHidrantesCamiones As String = "D51"
Private Const celdaHidrantesDirecta As String = "E51"
Private Const celdaBies As String = "E53"
Private Const celdaColumnaSeca As String = "A55"
Private Const celdaSistemasFijos As String = "B55"
Private Const celdaScteh As String = "C55"


Private Const celdaAlturaEvacuacion As String = "B7"
Private Const celdaFachadaAccesible As String = "B8"
Private Const celdaViabilidad As String = "F7"
Private Const celdaSuperficieViable As String = "D7"
Private Const celdaResistenciaViable As String = "D8"
Private Const celdaResistencia As String = "F9"
Private Const celdaNumeroSalidas As String = "F29"


' ============================
'  PUNTO DE ENTRADA
' ============================
Public Sub GenerarHojaResumenVGP()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    EnsureLocalTemplate

    Dim wsOut As Worksheet
    Set wsOut = RebuildFromTemplate(OUTPUT_SHEET_NAME)

    If WorksheetExists(MAP_SHEET) Then
        ApplySheetMappings wsOut, MAP_SHEET
    Else
        ApplyStaticMappings wsOut
    End If

    CalcularViabilidadBloqueO21
    CalcularViabilidadBloqueO22

    Application.GoTo wsOut.Range("A1"), True

fin:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrH:
    MsgBox "GenerarHojaResumenVGP: " & Err.Description, vbExclamation
    Resume fin
End Sub

' ============================
'  PLANTILLA LOCAL (primera vez)
' ============================
Private Sub EnsureLocalTemplate()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOCAL_TEMPLATE_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        If ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
        Exit Sub
    End If

    If Not WorksheetExists(TEMPLATE_SHEET_NAME) Then
        Err.Raise vbObjectError + 101, , _
            "No encuentro la plantilla '" & TEMPLATE_SHEET_NAME & "'. " & _
            "AÃ±Ã¡dela al libro y vuelve a ejecutar."
    End If

    ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME).Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ws.Name = LOCAL_TEMPLATE_SHEET
    ws.Visible = xlSheetVeryHidden
End Sub

' ============================
'  DUPLICAR PLANTILLA
' ============================
Private Function RebuildFromTemplate(ByVal outName As String) As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(outName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ThisWorkbook.Worksheets(LOCAL_TEMPLATE_SHEET).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(LOCAL_TEMPLATE_SHEET).Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ws.Name = outName
    ThisWorkbook.Worksheets(LOCAL_TEMPLATE_SHEET).Visible = xlSheetVeryHidden
    Set RebuildFromTemplate = ws
End Function

' ============================
'  MAPEOS
' ============================
' 1) Mapeo por hoja _MAP_VGP (Campo | Origen | Destino)
Private Sub ApplySheetMappings(ByVal wsOut As Worksheet, ByVal mapSheetName As String)
    Dim wsMap As Worksheet
    Set wsMap = ThisWorkbook.Worksheets(mapSheetName)

    Dim lastRow As Long, r As Long
    Dim origenRef As String, destinoAddr As String, v As Variant
    lastRow = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).row

    For r = 2 To lastRow
        origenRef = Trim$(CStr(wsMap.Cells(r, "B").Value))
        destinoAddr = Trim$(CStr(wsMap.Cells(r, "C").Value))
        If Len(origenRef) > 0 And Len(destinoAddr) > 0 Then
            v = GetValueByRef(origenRef)
            WriteValue wsOut, destinoAddr, v
        End If
    Next r
End Sub

' 2) Mapeo interno en BLOQUES (evita lÃ­mite de continuaciones)
Private Sub ApplyStaticMappings(ByVal wsOut As Worksheet)
    Dim mapArr As Variant
    mapArr = GetStaticMappings()

    Dim i As Long, origenRef As String, destinoAddr As String, v As Variant
    For i = LBound(mapArr) To UBound(mapArr)
        origenRef = CStr(mapArr(i)(0))
        destinoAddr = CStr(mapArr(i)(1))
        If Len(origenRef) > 0 And Len(destinoAddr) > 0 Then
            v = GetValueByRef(origenRef)
            WriteValue wsOut, destinoAddr, v
        End If
    Next i
End Sub

' Devuelve el mapeo por defecto en BLOQUES + concatenaciÃ³n
Private Function GetStaticMappings() As Variant
    Dim b1 As Variant, b2 As Variant, b3 As Variant, b4 As Variant, b5 As Variant, tmp As Variant

    b1 = Array( _
        Array(INTERFAZ_SHEET & "!" & celdaSuperficieSector, "E5"), _
        Array(INTERFAZ_SHEET & "!" & celdaQs, "J5"), _
        Array(INTERFAZ_SHEET & "!" & celdaActividad, "O5"), _
        Array(INTERFAZ_SHEET & "!" & celdaTipo, "E6"), _
        Array(INTERFAZ_SHEET & "!" & celdaNri, "J6"), _
        Array(INTERFAZ_SHEET & "!" & celdaNri, "O6"), _
        Array(INTERFAZ_SHEET & "!" & celdaSuperficieViable, "B14"), _
        Array(INTERFAZ_SHEET & "!" & celdaRasante, "D14"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlag50FachadaAccesible, "F14"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagRociadores, "H14"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagSinLimiteTipoCSistemaFijo, "J14"), _
        Array(INTERFAZ_SHEET & "!" & celdaSuperficieMaximaCorregida, "N14"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagBaterias, "O16"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagEstructurasIndependientes, "O17"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlag1Bie, "O19"), _
        Array(INTERFAZ_SHEET & "!" & celdaResistenciaViable, "B26"), _
        Array(INTERFAZ_SHEET & "!" & celdaResistencia, "G26"), _
        Array(INTERFAZ_SHEET & "!" & celdaPuertasPasoPeatoneal, "L26"))

    b2 = Array( _
        Array(INTERFAZ_SHEET & "!" & celdaParedesYTechos, "B29"), _
        Array(INTERFAZ_SHEET & "!" & celdaSuelos, "J29"), _
        Array(INTERFAZ_SHEET & "!" & celdaElementosSeparadores, "B35"), _
        Array(INTERFAZ_SHEET & "!" & celdaElementosSeparadoresColindantes, "E35"), _
        Array(INTERFAZ_SHEET & "!" & celdaReaccionAlFuegoDeFachadas, "O35"), _
        Array(INTERFAZ_SHEET & "!" & celdaLongitudRecorridoUnaSalidaOSinAlternativa, "B39"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagSistemaFijo, "D39"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagStech, "E39"), _
        Array(INTERFAZ_SHEET & "!" & celdaFactor2SalidasSistemaFijoHnave8, "F39"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagSistemaFijo, "J39"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagStech, "K39"), _
        Array(INTERFAZ_SHEET & "!" & celdaFactor2SalidasSistemaFijoHnave8, "L39"), _
        Array(INTERFAZ_SHEET & "!" & celdaNri, "M39"))

    b3 = Array( _
        Array(INTERFAZ_SHEET & "!" & celdaEstructuraPortante, "B57"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagCubiertaLigeraIndependiente, "F57"), _
        Array(INTERFAZ_SHEET & "!" & celdaFlagCubiertaLigeraConSistemaFijo, "J57"), _
        Array(INTERFAZ_SHEET & "!" & celdaDeteccionAutomatica, "D62"), _
        Array(INTERFAZ_SHEET & "!" & celdaDeteccionManual, "E62"), _
        Array("=""Si""", "J62"), _
        Array(INTERFAZ_SHEET & "!" & celdaHidrantesCamiones, "H62"), _
        Array(INTERFAZ_SHEET & "!" & celdaHidrantesDirecta, "I62"), _
        Array(INTERFAZ_SHEET & "!" & celdaBies, "K62"), _
        Array(INTERFAZ_SHEET & "!" & celdaColumnaSeca, "L62"), _
        Array(INTERFAZ_SHEET & "!" & celdaSistemasFijos, "M62"), _
        Array(INTERFAZ_SHEET & "!" & celdaScteh, "N62"), _
        Array(INTERFAZ_SHEET & "!" & celdaLongitudRecorridoSalidaExistente, "P39"), _
        Array("=""Si""", "O62"), _
        Array("=""Si""", "P62") _
        )

    b4 = Array( _
        Array("MITAD(" & INTERFAZ_SHEET & "!" & celdaElementosSeparadores & ")", "G35"), _
        Array("MITAD(" & INTERFAZ_SHEET & "!" & celdaElementosSeparadores & ")", "J35"), _
        Array("MITAD(" & INTERFAZ_SHEET & "!" & celdaElementosSeparadores & ")", "L35"), _
        Array( _
        "((" & INTERFAZ_SHEET & "!" & celdaAlturaEvacuacion & ">15)*" & _
        "(" & INTERFAZ_SHEET & "!" & celdaFlag50FachadaAccesible & ">5)*" & _
        "(" & INTERFAZ_SHEET & "!" & celdaRasante & "=""Sobre rasante""))<>0", _
        "O18" _
        ), _
        Array( _
        "=IF(OR(" & INTERFAZ_SHEET & "!" & celdaDeteccionAutomatica & "=""Si""," & _
        INTERFAZ_SHEET & "!" & celdaDeteccionManual & "=""Si""),""Si"",""No"")", _
        "F62" _
        ), _
        Array(INTERFAZ_SHEET & "!" & celdaLongitudRecorrido2Salidas, "I39"), _
        Array(INTERFAZ_SHEET & "!" & celdaLongitudRecorrido2Salidas, "N39"))

    b5 = Array( _
        Array( _
        "=IF(OR('" & INTERFAZ_SHEET & "'!" & celdaHidrantesCamiones & "=""Si""," & _
        "'" & INTERFAZ_SHEET & "'!" & celdaHidrantesDirecta & "=""Si""," & _
        "'" & INTERFAZ_SHEET & "'!" & celdaBies & "=""Si""),""Si"",""No"")", _
        "G62" _
        ), _
        Array( _
        "=IF('" & INTERFAZ_SHEET & "'!" & celdaSuperficieSector & "<100,""VERDADERO"",""FALSO"")", _
        "O20" _
        ), _
Array( _
    "=EsViableNota5('" & INTERFAZ_SHEET & "'!" & celdaViabilidad & ")", "L14"), _
Array( _
    "=" & IIf(Val(CStr(ActiveSheet.Range(celdaNumeroSalidas).Value)) = 1, _
              "'" & INTERFAZ_SHEET & "'!" & ActiveSheet.Range(celdaLongitudRecorridoUnaSalidaOSinAlternativa).Address(False, False), _
              "'" & INTERFAZ_SHEET & "'!" & ActiveSheet.Range(celdaLongitudRecorridoCorregidaUnaSalidaOSinAlternativa).Address(False, False)), _
    "B39"), _
Array( _
    "=" & IIf(Val(CStr(ActiveSheet.Range(celdaNumeroSalidas).Value)) = 1, _
              "'" & INTERFAZ_SHEET & "'!" & ActiveSheet.Range(celdaLongitudRecorridoUnaSalidaOSinAlternativa).Address(False, False), _
              "'" & INTERFAZ_SHEET & "'!" & ActiveSheet.Range(celdaLongitudRecorridoCorregidaUnaSalidaOSinAlternativa).Address(False, False)), _
    "G39"))

        
    tmp = ConcatArrays(b1, b2)
    tmp = ConcatArrays(tmp, b3)
    tmp = ConcatArrays(tmp, b4)
    tmp = ConcatArrays(tmp, b5)
    GetStaticMappings = tmp
End Function

' ============================
'  UTILIDADES
' ============================
Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

' Lee "Hoja!A1" o "A1" (asume Interfaz). Si contiene + - * /, evalÃºa la expresiÃ³n en el
' contexto de la hoja Interfaz (p. ej. "F24/2" o "F30/F31").
' Lee "Hoja!A1" o "A1" (asume Interfaz).
' AdemÃ¡s, si la referencia parece una EXPRESIÃ“N (comparadores o lÃ³gicas),
' la evalÃºa en el contexto de la hoja Interfaz.
Private Function GetValueByRef(ByVal ref As String) As Variant
    Dim p As Long, sh As String, addr As String, expr As String

    ' 1) Â¿Viene ya como fÃ³rmula? (=...)
    If Left$(Trim$(ref), 1) = "=" Then
        On Error GoTo EH_FORM
        GetValueByRef = Application.Evaluate(ref)
        Exit Function
    End If

    ' 2) Â¿Tiene pinta de EXPRESIÃ“N? (comparadores o lÃ³gicas)
    If IsExpressionCandidate(ref) Then
        expr = Trim$(ref)
        If Left$(expr, 1) <> "=" Then expr = "=" & expr
        On Error GoTo EH_FORM
        GetValueByRef = Application.Evaluate(ref)
        Exit Function
    End If

    ' 3) Caso simple: "Hoja!A1" o "A1"
    p = InStr(1, ref, "!", vbTextCompare)
    If p > 0 Then
        sh = Left$(ref, p - 1)
        addr = Mid$(ref, p + 1)
    Else
        sh = INTERFAZ_SHEET
        addr = ref
    End If

    ' --- Casos especiales de funciones propias ---
    If Left$(UCase$(Trim$(ref)), 6) = "MITAD(" Then
        Dim exprArg As String
        exprArg = Mid$(ref, 7, Len(ref) - 7) ' extrae lo que estÃ¡ dentro de MITAD(...)
        Dim valor As Variant
        valor = ThisWorkbook.Worksheets(INTERFAZ_SHEET).Range(exprArg).Value
        GetValueByRef = MitadTextoNumero(valor)
        Exit Function
    End If

    On Error GoTo EH_READ
    GetValueByRef = ThisWorkbook.Worksheets(sh).Range(addr).Value
    Exit Function

EH_FORM:
    GetValueByRef = Empty
    Exit Function
EH_READ:
    GetValueByRef = Empty
End Function

' HeurÃ­stica sencilla para detectar expresiones
Private Function IsExpressionCandidate(ByVal s As String) As Boolean
    Dim t As String: t = UCase$(s)
    If InStr(t, ">") > 0 Then IsExpressionCandidate = True: Exit Function
    If InStr(t, "<") > 0 Then IsExpressionCandidate = True: Exit Function
    If InStr(t, "=") > 0 Then IsExpressionCandidate = True: Exit Function
    If InStr(t, "AND(") > 0 Or InStr(t, " OR(") > 0 Or InStr(t, " NOT(") > 0 Then
        IsExpressionCandidate = True: Exit Function
    End If
    ' Por compatibilidad con lo aritmÃ©tico que ya tenÃ­as:
    If InStr(t, "+") Or InStr(t, "-") Or InStr(t, "*") Or InStr(t, "/") Or InStr(t, "^") Then
        IsExpressionCandidate = True
    End If
End Function

' Devuelve la celda escribible si el destino es combinado
Private Function ResolveWritableTarget(ByVal ws As Worksheet, ByVal destAddr As String) As Range
    Dim r As Range
    Set r = ws.Range(destAddr)
    If r.MergeCells Then
        Set ResolveWritableTarget = r.MergeArea.Cells(1, 1)
    Else
        Set ResolveWritableTarget = r
    End If
End Function

' Escribe valor respetando celdas combinadas
Private Sub WriteValue(ByVal ws As Worksheet, ByVal destAddr As String, ByVal v As Variant, _
    Optional ByVal keepFormulaIfAny As Boolean = False)
    Dim tgt As Range
    Set tgt = ResolveWritableTarget(ws, destAddr)
    If keepFormulaIfAny And tgt.HasFormula Then Exit Sub
    tgt.Value = v
End Sub

' Concatena dos arrays de 1D de variantes
Private Function ConcatArrays(a As Variant, b As Variant) As Variant
    Dim res() As Variant, i As Long, n As Long, k As Long
    If IsEmpty(a) Then
        ConcatArrays = b: Exit Function
    End If
    n = (UBound(a) - LBound(a) + 1) + (UBound(b) - LBound(b) + 1)
    ReDim res(0 To n - 1)
For i = LBound(a) To UBound(a): res(k) = a(i): k = k + 1: Next
    For i = LBound(b) To UBound(b): res(k) = b(i): k = k + 1: Next
            ConcatArrays = res
        End Function

        Private Function MitadTextoNumero(ByVal v As Variant) As String
            Dim re As Object, m As Object, num As Double, texto As String
            Set re = CreateObject("VBScript.RegExp")
            re.Pattern = "(\d+([.,]\d+)?)"
            re.Global = False

            If IsError(v) Or IsEmpty(v) Then
                MitadTextoNumero = ""
                Exit Function
            End If

            texto = CStr(v)

            If re.Test(texto) Then
                Set m = re.Execute(texto)
                num = CDbl(Replace(m(0).Value, ",", "."))
                num = num / 2
                texto = Replace(texto, m(0).Value, Trim$(Format(num, "0.##")))
            End If

            MitadTextoNumero = texto
        End Function

Public Sub CalcularViabilidadBloqueO21()
    Dim cfg As String
    Dim sup As Double
    Dim limite As Double
    Dim okEstructura As Boolean, okDeteccion As Boolean, okHumos As Boolean
    
    On Error GoTo ErrH
    
    '--- Lecturas desde Interfaz
    cfg = UCase$(Trim$(CStr(Worksheets(INTERFAZ_SHEET).Range("B2").Value))) ' configBuscado
    sup = ToDouble(Worksheets(INTERFAZ_SHEET).Range("F3").Value)             ' celdaSuperficieSector (nÃºmero)
    
    okEstructura = ToBool(Worksheets(INTERFAZ_SHEET).Range("B18").Value)     ' flagEstructura
    okDeteccion = ToBool(Worksheets(INTERFAZ_SHEET).Range("B17").Value)      ' flagDetecciÃ³n
    okHumos = ToBool(Worksheets(INTERFAZ_SHEET).Range("B20").Value)          ' flagHumos
    
    '--- Si no estÃ¡n los tres flags, directamente NO viable
    If Not (okEstructura And okDeteccion And okHumos) Then
        Worksheets(OUTPUT_SHEET_NAME).Range("O21").Value = False
        Exit Sub
    End If
    
    '--- LÃ­mite por configuraciÃ³n
    Select Case cfg
        Case "AV": limite = 300
        Case "AH": limite = 1500
        Case "B":  limite = 3000
        Case Else
            ' Config no reconocida -> no viable
            Worksheets(OUTPUT_SHEET_NAME).Range("O21").Value = False
            Exit Sub
    End Select
    
    '--- ComprobaciÃ³n de superficie y escritura del booleano en O21
    Worksheets(OUTPUT_SHEET_NAME).Range("O21").Value = (sup <= limite)
    Exit Sub
    
ErrH:
    ' Ante cualquier problema, deja FALSO para no dar falsos positivos
    Worksheets(OUTPUT_SHEET_NAME).Range("O21").Value = False
End Sub

Public Sub CalcularViabilidadBloqueO22()
    Dim cfg As String
    Dim sup As Double
    Dim okEstructura As Boolean, okDeteccion As Boolean, okHumos As Boolean, okSeguridad As Boolean
    Dim esValido As Boolean
    
    On Error GoTo ErrH
    
    '--- Entradas desde Interfaz
    cfg = UCase$(Trim$(CStr(Worksheets(INTERFAZ_SHEET).Range("B2").Value))) ' configBuscado ("AH" | "B" | ...)
    sup = ToDouble(Worksheets(INTERFAZ_SHEET).Range("F3").Value)             ' superficieBuscada (nÃºmero)
    
    okEstructura = ToBool(Worksheets(INTERFAZ_SHEET).Range("B18").Value)     ' flagEstructura
    okDeteccion = ToBool(Worksheets(INTERFAZ_SHEET).Range("B17").Value)      ' flagDetecciÃ³n
    okHumos = ToBool(Worksheets(INTERFAZ_SHEET).Range("B20").Value)          ' flagHumos
    okSeguridad = ToBool(Worksheets(INTERFAZ_SHEET).Range("B22").Value)      ' flagSeguridad
    
    '--- Si no estÃ¡n los tres flags base, directamente FALSO
    If Not (okEstructura And okDeteccion And okHumos) Then
        Worksheets(OUTPUT_SHEET_NAME).Range("O22").Value = False
        Exit Sub
    End If
    
    '--- LÃ³gica del bloque 3
    esValido = False
    If okSeguridad Then
        Select Case cfg
            Case "AH"
                esValido = (sup <= 5000)
            Case "B"
                esValido = (sup <= 6000)
            Case Else
                esValido = False
        End Select
    End If
    
    '--- Escribir resultado en Tabla resumen!O22 (VERDADERO/FALSO)
    Worksheets(OUTPUT_SHEET_NAME).Range("O22").Value = esValido
    Exit Sub

ErrH:
    Worksheets(OUTPUT_SHEET_NAME).Range("O22").Value = False
End Sub

Private Function ToBool(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        ToBool = v
        Exit Function
    End If
    Dim t As String
    t = UCase$(Trim$(CStr(v)))
    ToBool = (t = "SI" Or t = "SÃ" Or t = "TRUE" Or t = "VERDADERO" Or t = "1" Or t = "X")
End Function

Private Function ToDouble(ByVal v As Variant) As Double
    ' Convierte nÃºmeros con coma o punto
    Dim s As String
    s = Trim$(CStr(v))
    If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then s = Replace(s, ",", ".")
    ToDouble = Val(s)
End Function
' Devuelve TRUE si la celda contiene (en cualquier formato) el texto "ADMITIDO CON NOTA 5"
Public Function EsViableNota5(r As Range) As Boolean
    On Error GoTo fin
    Dim s As String
    s = UCase$(Trim$(CStr(r.Value2)))
    EsViableNota5 = (InStr(1, s, "ADMITIDO CON NOTA 5", vbTextCompare) > 0)
    Exit Function
fin:
    EsViableNota5 = False
End Function

' Devuelve TRUE si el n�mero de salidas es 1, aunque est� como "1", "1,00", "1 salida", etc.
Public Function EsUnaSalida(r As Range) As Boolean
    On Error GoTo fin
    Dim v As Variant, s As String, i As Long, ch As String, numTxt As String

    v = r.Value2
    If IsNumeric(v) Then
        EsUnaSalida = (CLng(v) = 1)
        Exit Function
    End If

    s = Trim$(CStr(v))
    ' Extrae el primer d�gito/entero que aparezca
    numTxt = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Then
            numTxt = numTxt & ch
        ElseIf numTxt <> "" Then
            Exit For
        End If
    Next i

    If numTxt <> "" Then
        EsUnaSalida = (CLng(numTxt) = 1)
    Else
        ' Fallback: mira primer car�cter por si es "1"
        EsUnaSalida = (Left$(s, 1) = "1")
    End If
    Exit Function
fin:
    EsUnaSalida = False
End Function
