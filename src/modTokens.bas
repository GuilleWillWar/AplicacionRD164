Attribute VB_Name = "modTokens"
Option Explicit

'============================
'   GESTIÓN DE TOKENS UNIFICADA
'============================

' === Tipo de token configurable ===
Public Type TToken
    Pattern As String     ' Patrón de búsqueda (Regex o literal escapado)
    Replacement As String ' Valor que se insertará
    IsRegex As Boolean    ' TRUE = Pattern es expresión regular; FALSE = literal
    escapeRTF As Boolean  ' TRUE = escapar Replacement a RTF (para construir RTF)
End Type


' ========================================================
' CARGA TOKENS (SCALAR y TXT) DESDE HOJA "TOKENS"
' ========================================================
'
' Estructura de la hoja TOKENS (fila 1 cabeceras):
'   A: TIPO          -> "SCALAR" o "TXT"
'   B: TOKEN_ID      -> nombre del token (p.ej. {CLIENTE} o DECRIPCION_DETECCION)
'   C: ORIGEN        -> para SCALAR: texto o referencia (=Interfaz!E56, nombre definido, etc.)
'   D: CONFIG        -> para TXT: AH / AV / B / C / D (o "*")
'   E: NRI           -> para TXT: Bajo 1 ... Alto 8 (o "*")
'   F: TEXTO         -> texto del párrafo (solo TXT)
'   G: PRIORIDAD     -> número (mayor = gana si no hay MULTI)
'   H: MULTI         -> TRUE/FALSE (concatena si TRUE)
'   I: ESCAPE_RTF    -> TRUE/FALSE (por defecto TRUE)
'   J: ES_REGEX      -> TRUE/FALSE (solo SCALAR, por defecto FALSE)
'   K: ACTIVO        -> TRUE/FALSE (por defecto TRUE)
'
' Nombres definidos esperados:
'   CONFIGURACION, NRI
'
Public Function CargarTokens(Optional ByVal wb As Workbook) As TToken()
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim conf$, nri$
    Dim scalars() As TToken, idxS As Long
    Dim dictTXT As Object: Set dictTXT = CreateObject("Scripting.Dictionary")

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets("TOKENS")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Function

    ' Variables CONFIGURACION y NRI desde nombres definidos
    conf = ValorDeNombre(wb, "CONFIGURACION", "")
    nri = ValorDeNombre(wb, "NRI", "")

    ' --- Recorremos filas ---
    For i = 2 To lastRow
        Dim tipo$, tokenId$, origen$, cfg$, nr$, texto$
        Dim prio As Variant, multi As Variant, esc As Variant, esrx As Variant, activo As Variant

        tipo = UCase$(Trim$(CStr(ws.Cells(i, "A").Value)))
        tokenId = Trim$(CStr(ws.Cells(i, "B").Value))
        origen = CStr(ws.Cells(i, "C").Value)
        cfg = Trim$(CStr(ws.Cells(i, "D").Value))
        nr = Trim$(CStr(ws.Cells(i, "E").Value))
        texto = CStr(ws.Cells(i, "F").Value)
        prio = ws.Cells(i, "G").Value
        multi = ws.Cells(i, "H").Value
        esc = ws.Cells(i, "I").Value
        esrx = ws.Cells(i, "J").Value
        activo = ws.Cells(i, "K").Value

        If Len(tokenId) = 0 Then GoTo siguiente
        If CBool(IIf(IsEmpty(activo) Or activo = "", True, activo)) = False Then GoTo siguiente

        Select Case tipo
        Case "SCALAR", ""
            ' --- Token simple ---
            idxS = idxS + 1
            If (Not Not scalars) = 0 Then
                ReDim scalars(1 To 1)
            Else
                ReDim Preserve scalars(1 To idxS)
            End If

            scalars(idxS).IsRegex = CBool(IIf(IsEmpty(esrx) Or esrx = "", False, esrx))
            scalars(idxS).escapeRTF = CBool(IIf(IsEmpty(esc) Or esc = "", True, esc))

            If scalars(idxS).IsRegex Then
                scalars(idxS).Pattern = tokenId
            Else
                scalars(idxS).Pattern = RegexEscapeLiteral(tokenId)
            End If

            scalars(idxS).Replacement = EvaluarOrigenToken(origen, wb)

        Case "TXT"
            ' --- Token de texto condicionado ---
            Dim rec As Variant
            rec = Array(tokenId, cfg, nr, texto, prio, multi, esc)
            If Not dictTXT.Exists(tokenId) Then
                dictTXT.add tokenId, Array(rec)
            Else
                dictTXT(tokenId) = AppendRec(dictTXT(tokenId), rec)
            End If
        End Select
siguiente:
    Next i

    ' --- Resolver los TXT ---
    Dim txtTokens() As TToken, idxT As Long, k As Variant
    If dictTXT.Count > 0 Then
        ReDim txtTokens(1 To dictTXT.Count): idxT = 0
        For Each k In dictTXT.Keys
            Dim textoFinal As String, escapeTxt As Boolean
            textoFinal = SeleccionarTextoPorDosReglas(dictTXT(k), conf, nri, escapeTxt)
            idxT = idxT + 1
            txtTokens(idxT).Pattern = RegexEscapeLiteral("{{TXT:" & CStr(k) & "}}")
            txtTokens(idxT).IsRegex = False
            txtTokens(idxT).escapeRTF = escapeTxt
            txtTokens(idxT).Replacement = textoFinal
        Next k
    End If

    ' --- Combinar ---
    CargarTokens = CombinarTokens(scalars, txtTokens)
End Function


' ========================================================
' FUNCIONES DE APOYO
' ========================================================

Private Function SeleccionarTextoPorDosReglas(ByVal recs As Variant, _
                                              ByVal conf As String, _
                                              ByVal nri As String, _
                                              ByRef escapeRTF As Boolean) As String
    Dim i&, anyMulti As Boolean
    Dim elegibles As Object: Set elegibles = CreateObject("System.Collections.ArrayList")
    escapeRTF = True

    Dim confN$, nriN$
    confN = Trim$(conf): nriN = Trim$(nri)

    For i = LBound(recs) To UBound(recs)
        Dim cfg$, nr$, txt$, prio As Double, multi As Boolean, esc As Boolean
        cfg = recs(i)(1): nr = recs(i)(2): txt = CStr(recs(i)(3))
        prio = CDbl(Val(recs(i)(4)))
        multi = CBool(IIf(recs(i)(5) = "", False, recs(i)(5)))
        esc = CBool(IIf(recs(i)(6) = "", True, recs(i)(6)))

        If IgualOCumplePatron(confN, cfg) And IgualOCumplePatron(nriN, nr) Then
            elegibles.add Array(cfg, nr, txt, prio, multi, esc)
            If multi Then anyMulti = True
            If esc = False Then escapeRTF = False
        End If
    Next i

    If elegibles.Count = 0 Then
        SeleccionarTextoPorDosReglas = ""
        Exit Function
    End If

    ' MULTI = TRUE ? concatenar todas (orden por PRIORIDAD)
    If anyMulti Then
        Dim j&, k&, maxIdx&, out As String
        For j = 0 To elegibles.Count - 2
            maxIdx = j
            For k = j + 1 To elegibles.Count - 1
                If CDbl(elegibles(k)(3)) > CDbl(elegibles(maxIdx)(3)) Then maxIdx = k
            Next
            If maxIdx <> j Then
                Dim tmp: tmp = elegibles(j): elegibles(j) = elegibles(maxIdx): elegibles(maxIdx) = tmp
            End If
        Next
        For j = 0 To elegibles.Count - 1
            If out <> "" Then out = out & vbCrLf
            out = out & CStr(elegibles(j)(2))
        Next
        SeleccionarTextoPorDosReglas = out
        Exit Function
    End If

    ' Sin MULTI ? escoger el de mayor prioridad
    Dim bestIdx As Long: bestIdx = 0
    For i = 1 To elegibles.Count - 1
        If CDbl(elegibles(i)(3)) > CDbl(elegibles(bestIdx)(3)) Then bestIdx = i
    Next
    SeleccionarTextoPorDosReglas = CStr(elegibles(bestIdx)(2))
End Function


Private Function IgualOCumplePatron(ByVal valor As String, ByVal patron As String) As Boolean
    Dim v$, p$
    v = UCase$(Trim$(valor))
    p = UCase$(Trim$(patron))
    If p = "*" Or p = "" Then IgualOCumplePatron = True Else IgualOCumplePatron = (v = p)
End Function

Private Function AppendRec(arr As Variant, rec As Variant) As Variant
    Dim n&, tmp()
    If IsEmpty(arr) Then
        ReDim tmp(0 To 0)
        tmp(0) = rec
    Else
        n = UBound(arr) - LBound(arr) + 1
        ReDim tmp(0 To n)
        Dim i&
        For i = 0 To n - 1: tmp(i) = arr(i): Next
        tmp(n) = rec
    End If
    AppendRec = tmp
End Function

Private Function ValorDeNombre(ByVal wb As Workbook, ByVal nm As String, Optional ByVal def As String = "") As String
    On Error GoTo fallo
    Dim v As Variant
    If TieneNombreDefinido(wb, nm) Then
        v = Application.Evaluate("=" & nm)
        If Not IsError(v) And Not IsEmpty(v) Then
            ValorDeNombre = Trim$(CStr(v))
            Exit Function
        End If
    End If
fallo:
    ValorDeNombre = def
End Function

Private Function TieneNombreDefinido(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim n As Name
    On Error Resume Next
    For Each n In wb.Names
        If StrComp(n.Name, nm, vbTextCompare) = 0 Then TieneNombreDefinido = True: Exit Function
        If InStr(1, n.Name, wb.Name & "!", vbTextCompare) > 0 Then
            If StrComp(Split(n.Name, "!")(1), nm, vbTextCompare) = 0 Then TieneNombreDefinido = True: Exit Function
        End If
    Next n
End Function

Private Function RegexEscapeLiteral(ByVal s As String) As String
    Dim ch As String, i As Long, out As String
    Const META As String = "\.^$|?*+()[]{}"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(1, META, ch, vbBinaryCompare) > 0 Then out = out & "\" & ch Else out = out & ch
    Next i
    RegexEscapeLiteral = out
End Function

Private Function CombinarTokens(ByRef a() As TToken, ByRef b() As TToken) As TToken()
    Dim out() As TToken
    Dim nA As Long, nB As Long, i As Long, idx As Long

    If (Not Not a) = 0 Then nA = 0 Else nA = UBound(a) - LBound(a) + 1
    If (Not Not b) = 0 Then nB = 0 Else nB = UBound(b) - LBound(b) + 1
    If nA + nB = 0 Then Exit Function

    ReDim out(1 To nA + nB): idx = 0
    If nA > 0 Then For i = LBound(a) To UBound(a): idx = idx + 1: out(idx) = a(i): Next i
    If nB > 0 Then For i = LBound(b) To UBound(b): idx = idx + 1: out(idx) = b(i): Next i
    CombinarTokens = out
End Function


' ========================================================
' REUTILIZABLES EXISTENTES (EscRTF y Marcadores)
' ========================================================

Private Function EsMarcadorTabla(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    EsMarcadorTabla = (Left$(t, 8) = "@@TABLA:" And Right$(t, 2) = "@@")
End Function

Public Function ResolverTokensEnLinea(ByVal linea As String, ByRef tokens() As TToken) As String
    Dim i As Long, re As Object, repl As String

    If EsMarcadorTabla(linea) Then
        ResolverTokensEnLinea = linea
        Exit Function
    End If

    ResolverTokensEnLinea = linea
    If (Not Not tokens) = 0 Then Exit Function

    For i = LBound(tokens) To UBound(tokens)
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = True
        re.Pattern = tokens(i).Pattern

        repl = tokens(i).Replacement
        If tokens(i).escapeRTF Then repl = EscRTF(repl)

        On Error Resume Next
        ResolverTokensEnLinea = re.Replace(ResolverTokensEnLinea, repl)
        On Error GoTo 0
    Next i
End Function

' === Evalúa el campo ORIGEN de un token SCALAR ===
' Reglas:
' - Si empieza por "=", se evalúa con Application.Evaluate.
' - Si contiene "!" (Hoja!A1) y no empieza por "=", se evalúa como "=" & origen.
' - Si es un NOMBRE DEFINIDO del libro, se evalúa como "=" & origen.
' - En cualquier otro caso, se devuelve literal.
Public Function EvaluarOrigenToken(ByVal origen As String, ByVal wb As Workbook) As String
    Dim expr As String, v As Variant, prev As Workbook

    On Error GoTo devolver_literal

    origen = Trim$(CStr(origen))
    If Len(origen) = 0 Then GoTo devolver_literal

    If Left$(origen, 1) = "=" Then
        expr = origen
    ElseIf InStr(1, origen, "!", vbTextCompare) > 0 Or TieneNombreDefinido(wb, origen) Then
        expr = "=" & origen
    Else
        GoTo devolver_literal
    End If

    Set prev = Application.ActiveWorkbook
    If Not wb Is Nothing Then wb.Activate
    v = Application.Evaluate(expr)
    If Not prev Is Nothing Then prev.Activate

    If IsError(v) Or IsEmpty(v) Then
        EvaluarOrigenToken = ""
    Else
        EvaluarOrigenToken = CStr(v)
    End If
    Exit Function

devolver_literal:
    EvaluarOrigenToken = CStr(origen)
End Function

