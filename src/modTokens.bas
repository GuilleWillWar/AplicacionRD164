Attribute VB_Name = "modTokens"
Option Explicit

' === Tipo de token configurable desde hoja TOKENS/TOKEN ===
Public Type TToken
    Pattern As String     ' Patrón de búsqueda (Regex o literal escapado)
    Replacement As String ' Valor que se insertará
    IsRegex As Boolean    ' TRUE = Pattern es expresión regular; FALSE = literal
    EscapeRTF As Boolean  ' TRUE = escapar Replacement a RTF (para construir RTF)
End Type

' Carga tokens desde la hoja "TOKENS" (preferido) o "TOKEN" del ThisWorkbook.
' Estructura de la hoja (fila 1 cabeceras):
'   A: TOKEN            -> lo que escribes en DOC_Config (p.ej. {{RAZON_SOCIAL}})
'   B: ORIGEN           -> texto o referencia (=Interfaz!E56, nombres definidos, etc.)
'   C: ES_REGEX         -> TRUE/FALSE (si A ya es patrón regex)
'   D: ESCAPE_RTF       -> TRUE/FALSE (por defecto TRUE)
Public Function CargarTokens(Optional ByVal wb As Workbook) As TToken()
    Dim ws As Worksheet, lastRow As Long, i As Long, idx As Long
    Dim tmp() As TToken
    Dim tok As String, src As String, isrx As Variant, esc As Variant

    If wb Is Nothing Then Set wb = ThisWorkbook

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("TOKENS")
    If ws Is Nothing Then Set ws = wb.Worksheets("TOKEN")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function ' devolverá array sin dimensionar

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Function

    ReDim tmp(1 To lastRow - 1)
    idx = 0

    For i = 2 To lastRow
        tok = Trim$(CStr(ws.Cells(i, 1).Value))
        src = Trim$(CStr(ws.Cells(i, 2).Value))
        isrx = ws.Cells(i, 3).Value
        esc = ws.Cells(i, 4).Value

        If Len(tok) > 0 Then
            idx = idx + 1
            tmp(idx).Pattern = tok
            tmp(idx).IsRegex = CBool(IIf(IsEmpty(isrx) Or isrx = "", False, isrx))
            tmp(idx).EscapeRTF = CBool(IIf(IsEmpty(esc) Or esc = "", True, esc))
            tmp(idx).Replacement = EvaluarOrigenToken(src, wb)
            ' Si no es regex, transformar token a literal-regex-seguro
            If tmp(idx).IsRegex = False Then
                tmp(idx).Pattern = RegexEscapeLiteral(tmp(idx).Pattern)
            End If
        End If
    Next i

    If idx = 0 Then Exit Function
    ReDim Preserve tmp(1 To idx)
    CargarTokens = tmp
End Function

' Evalúa el origen:
' - Si empieza por "=", se evalúa con Application.Evaluate.
' - Si tiene "!" (Interfaz!E56) y no empieza por "=", se convierte a "="+origen.
' - Si coincide con nombre definido, también se evalúa.
' - Si no, se toma como texto literal.
Private Function EvaluarOrigenToken(ByVal origen As String, ByVal wb As Workbook) As String
    Dim expr As String, v As Variant
    On Error GoTo literal
    If Len(Trim$(origen)) = 0 Then GoTo literal

    If Left$(origen, 1) = "=" Then
        expr = origen
    ElseIf InStr(1, origen, "!", vbTextCompare) > 0 Or TieneNombreDefinido(wb, origen) Then
        expr = "=" & origen
    Else
        GoTo literal
    End If

    Dim prev As Workbook
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

literal:
    EvaluarOrigenToken = CStr(origen)
End Function

Private Function TieneNombreDefinido(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim n As Name
    On Error Resume Next
    For Each n In wb.names
        If StrComp(n.Name, nm, vbTextCompare) = 0 Then TieneNombreDefinido = True: Exit Function
        If InStr(1, n.Name, wb.Name & "!", vbTextCompare) > 0 Then
            If StrComp(Split(n.Name, "!")(1), nm, vbTextCompare) = 0 Then TieneNombreDefinido = True: Exit Function
        End If
    Next n
End Function

' Escapa un literal para usarlo como patrón Regex (matchear el token tal cual)
Private Function RegexEscapeLiteral(ByVal s As String) As String
    Dim ch As String, i As Long, out As String
    Const META As String = "\.^$|?*+()[]{}"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(1, META, ch, vbBinaryCompare) > 0 Then out = out & "\" & ch Else out = out & ch
    Next i
    RegexEscapeLiteral = out
End Function

' Devuelve TRUE si la línea es un marcador de tabla @@TABLA:...@@ (se deja intacta)
Private Function EsMarcadorTabla(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    EsMarcadorTabla = (Left$(t, 8) = "@@TABLA:" And Right$(t, 2) = "@@")
End Function

' Aplica todos los tokens a una línea de texto.
' - Si la línea es marcador de tabla @@TABLA:...@@, se devuelve tal cual.
' - Si EscapeRTF=True, se escapa Replacement con EscRTF (reutiliza función de Modulo2).
Public Function ResolverTokensEnLinea(ByVal linea As String, ByRef tokens() As TToken) As String
    Dim i As Long
    Dim re As Object
    Dim repl As String

    If EsMarcadorTabla(linea) Then
        ResolverTokensEnLinea = linea
        Exit Function
    End If

    ResolverTokensEnLinea = linea
    If (Not Not tokens) = 0 Then Exit Function ' array sin dimensionar

    For i = LBound(tokens) To UBound(tokens)
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = True
        re.Pattern = tokens(i).Pattern

        repl = tokens(i).Replacement
        If tokens(i).EscapeRTF Then repl = EscRTF(repl)

        On Error Resume Next
        ResolverTokensEnLinea = re.Replace(ResolverTokensEnLinea, repl)
        On Error GoTo 0
    Next i
End Function

' ===== Utilidades para ocultar/mostrar hoja TOKENS =====
Public Sub OcultarHojaTokens(Optional ByVal wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    If Not wb.Worksheets("TOKENS") Is Nothing Then
        wb.Worksheets("TOKENS").Visible = xlSheetVeryHidden
    ElseIf Not wb.Worksheets("TOKEN") Is Nothing Then
        wb.Worksheets("TOKEN").Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0
End Sub

Public Sub MostrarHojaTokens(Optional ByVal wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    If Not wb.Worksheets("TOKENS") Is Nothing Then
        wb.Worksheets("TOKENS").Visible = xlSheetVisible
    ElseIf Not wb.Worksheets("TOKEN") Is Nothing Then
        wb.Worksheets("TOKEN").Visible = xlSheetVisible
    End If
    On Error GoTo 0
End Sub

