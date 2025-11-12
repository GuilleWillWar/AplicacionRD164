Attribute VB_Name = "ExportarMemoria"
Option Explicit

'==========================
'   GENERACIÓN DE MEMORIA
'==========================

Public Sub GenerarMemoria()
    On Error GoTo EH

    Const C_RUTA_ESQUEMA As String = "\Recursos\EsquemaCuna.xlsx"
    Const C_NOMBRE_HOJA As String = "DOC_Config"

    If Not ValidarInterfaz() Then Exit Sub

    Dim rutaEsq$, rtf$, tieneDatos As Boolean
    rutaEsq = RutaAbs(ThisWorkbook.path, C_RUTA_ESQUEMA)

    ' === NUEVO: cargar tokens desde la hoja TOKENS/TOKEN del ThisWorkbook ===
    Dim tokens() As TToken
    tokens = CargarTokens(ThisWorkbook) ' Si no existe la hoja, tokens quedará sin dimensionar

    If Dir$(rutaEsq, vbNormal Or vbHidden Or vbSystem) <> "" Then
        Dim wb As Workbook, ws As Worksheet, data As Variant
        Set wb = Workbooks.Open(fileName:=rutaEsq, ReadOnly:=True)

        On Error Resume Next
        Set ws = wb.Worksheets(C_NOMBRE_HOJA)
        On Error GoTo EH

        If Not ws Is Nothing Then
            If WorksheetFunction.CountA(ws.Cells) > 0 Then
                data = SafeRangeTo2D(ws.UsedRange)
                If UBound(data, 1) >= 2 Then
                    ' === PASAMOS TOKENS AL CONSTRUCTOR ===
                    rtf = ConstruirRTFDesdeDOCConfig(data, tokens)
                    tieneDatos = True
                End If
            End If
        End If

        wb.Close False
    End If

    If Not tieneDatos Then
        rtf = ConstruirRTFFallback()
    End If

    ' === 1) Guardar RTF temporal en %TEMP% ===
    Dim tempOut$, desktopOut$, nombreBase$, nombreRTF$, rutaRTF$, rutaWordOut$
    tempOut = TempPath()
    desktopOut = DesktopPath()

    nombreBase = LimpiarNombreArchivo("PXL_" & _
                 CStr(ThisWorkbook.Worksheets("Interfaz").Range("E56").Value) & "_" & _
                 CStr(ThisWorkbook.Worksheets("Interfaz").Range("E57").Value) & "_" & _
                 CStr(ThisWorkbook.Worksheets("Interfaz").Range("E58").Value))
    If nombreBase = "" Then nombreBase = "Memoria_SinDatos"

    nombreRTF = nombreBase & ".rtf"
    rutaRTF = tempOut & nombreRTF
    SaveTextFile rutaRTF, rtf

    ' === 2) Abrir en Word, sustituir marcadores y guardar DOCX en Escritorio ===
    Dim wdApp As Object, wdDoc As Object, wordYaAbierto As Boolean

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    wordYaAbierto = Not wdApp Is Nothing
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo EH

    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(rutaRTF)
    
    ' 1) Crear/ajustar estilos corporativos
    ConfigurarEstilosPXL wdDoc

    ' 2) Mapear marcadores [[H1]]/[[H2]]/[[H3]]/[[P]] a estilos
    AplicarEstilosPorMarcadores wdDoc

    ' Sustituye todos los @@TABLA:...@@
    Call SustituirMarcadoresTablas(wdDoc)
    
    'Sustituye las imagenes
    Call SustituirMarcadoresImagenes(wdDoc)

    ' --- OPCIONAL: si alguna vez necesitas sustituir tokens en tablas/headers/footers ---
    'Call SustituirTokensEnTablasDeWord(wdDoc, tokens)
    'Call SustituirTokensEnCuerpoYSecciones(wdDoc, tokens)
    
    Call AplicarFormatoMemoria(wdDoc)

    rutaWordOut = EvitarColisiones(desktopOut, nombreBase & ".docx")
    wdDoc.SaveAs2 desktopOut & rutaWordOut, 16  ' wdFormatXMLDocument
    wdDoc.Close False

    If Not wordYaAbierto Then wdApp.Quit

    ' === 3) Eliminar el RTF temporal ===
    Call DeleteIfExists(rutaRTF)

    MsgBox "Memoria generada correctamente:" & vbCrLf & _
           desktopOut & rutaWordOut, vbInformation
    Exit Sub

EH:
    MsgBox "GenerarMemoria -> " & Err.Number & ": " & Err.Description, vbCritical
End Sub


' ========= CONSTRUCCIÓN RTF (con tokens) =========
Private Function ConstruirRTFDesdeDOCConfig(ByVal data As Variant, ByRef tokens() As TToken) As String

    Dim iTipo&, iEstilo&, iTexto&, iSalto&
    iTipo = FindHeaderIndex(data, "Tipo")
    iEstilo = FindHeaderIndex(data, "Estilo")
    iTexto = FindHeaderIndex(data, "Texto")
    iSalto = FindHeaderIndex(data, "SaltoPagina")

    Dim rtf$, r&, tipo$, estilo$, linea$, salto$
    rtf = RTF_Header()

    Dim nFilas&: nFilas = UBound(data, 1)
    For r = 2 To nFilas
        tipo = UCase$(NzSafe(GetCell(data, r, iTipo)))
        estilo = NzSafe(GetCell(data, r, iEstilo))
        linea = NzSafe(GetCell(data, r, iTexto))
        If iSalto > 0 Then salto = UCase$(NzSafe(GetCell(data, r, iSalto))) Else salto = ""

        ' Asignar estilo por tipo
If estilo = "" Then
    Select Case tipo
        Case "TITULO", "TÍTULO": estilo = "Titulo 1"
        Case "AVISO" ' << NUEVO: si Tipo = AVISO, por defecto rojo
            estilo = "PXL_Rojo"
        Case Else
            estilo = "Normal"
    End Select
End If

        ' === NUEVO: sustituir variables usando el mapa TOKENS ===
        On Error Resume Next
        linea = ResolverTokensEnLinea(linea, tokens)
        linea = ReemplazarTokensIMG(linea)
        On Error GoTo 0
        
        If EtiquetaEsTablaCandidata(linea) Then
            ' Insertamos un marcador que luego reemplazaremos en Word por la tabla real
            rtf = rtf & RTF_Parrafo("@@TABLA:" & linea & "@@", "Normal")
        Else
            rtf = rtf & RTF_Parrafo(linea, estilo)
        End If

        If salto = "SI" Or salto = "YES" Then rtf = rtf & "\page" & vbCrLf
    Next r

    rtf = rtf & "}"
    ConstruirRTFDesdeDOCConfig = rtf
End Function


' ========= FALLBACK =========
Private Function ConstruirRTFFallback() As String
    Dim rtf$
    rtf = RTF_Header()
    rtf = rtf & RTF_Parrafo("No se pudo leer 'DOC_Config'. Contenido de respaldo:", "Titulo 2")
    rtf = rtf & RTF_Parrafo("• Título 1", "Titulo 3")
    rtf = rtf & RTF_Parrafo("Texto de ejemplo.", "Normal")
    rtf = rtf & "\page" & vbCrLf
    rtf = rtf & RTF_Parrafo("• Título 2", "Titulo 3")
    rtf = rtf & RTF_Parrafo("Texto de respaldo.", "Normal") & "}"
    ConstruirRTFFallback = rtf
End Function


' ========= HELPERS =========
Private Function SafeRangeTo2D(rg As Range) As Variant
    On Error GoTo EH
    Dim v As Variant: v = rg.Value
    If IsArray(v) Then
        SafeRangeTo2D = v
    Else
        Dim tmp(1 To 1, 1 To 1) As Variant
        tmp(1, 1) = v
        SafeRangeTo2D = tmp
    End If
    Exit Function
EH:
    Dim tmp2(1 To 1, 1 To 1) As Variant
    tmp2(1, 1) = ""
    SafeRangeTo2D = tmp2
End Function

Private Function GetCell(ByVal arr As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c <= 0 Then GetCell = "": Exit Function
    GetCell = arr(r, c)
End Function

Private Function FindHeaderIndex(ByVal arr As Variant, ByVal headerName As String) As Long
    On Error GoTo Salir
    Dim c&, maxC&: maxC = UBound(arr, 2)
    For c = 1 To maxC
        If StrComp(CStr(arr(1, c)), headerName, vbTextCompare) = 0 Then
            FindHeaderIndex = c: Exit Function
        End If
    Next c
Salir:
End Function

Private Function RTF_Header() As String
    RTF_Header = "{\rtf1\ansi\deff0" & _
                 "{\fonttbl{\f0 Calibri;}}" & _
                 "{\colortbl;\red0\green0\blue0;}" & _
                 "\viewkind4\uc1\pard\fs22" & vbCrLf
End Function

Private Function MkFromEstilo(ByVal estilo As String) As String
    Dim e$: e = UCase$(Replace(estilo, "Í", "I"))
    e = Replace(e, "Á", "A")
    e = Replace(e, "É", "E")
    e = Replace(e, "Ó", "O")
    e = Replace(e, "Ú", "U")

    Select Case e
        Case "PXL_TÍTULO 1", "PXL_TITULO 1", "TITULO 1", "TÍTULO 1", "TITULO1"
            MkFromEstilo = "[[H1]]"
        Case "PXL_TÍTULO 2", "PXL_TITULO 2", "TITULO 2", "TÍTULO 2", "TITULO2"
            MkFromEstilo = "[[H2]]"
        Case "PXL_TÍTULO 3", "PXL_TITULO 3", "TITULO 3", "TÍTULO 3", "TITULO3"
            MkFromEstilo = "[[H3]]"
        Case "PXL_ROJO", "ROJO"
            MkFromEstilo = "[[R]]"   ' << NUEVO: marcador para PXL_Rojo
        Case Else
            MkFromEstilo = "[[P]]"
    End Select
End Function

Private Function RTF_Parrafo(ByVal txt As String, ByVal estilo As String) As String
    ' Prependemos un marcador visible (que luego quitaremos en Word)
    Dim mk$: mk = MkFromEstilo(estilo)
    Dim run$: run = "\pard"
    ' Tamaños razonables, el estilo final NO dependerá de esto
    Select Case mk
        Case "[[H1]]": run = run & "\sa200\fs36\b "
        Case "[[H2]]": run = run & "\sa160\fs28\b "
        Case "[[H3]]": run = run & "\sa120\fs24\b "
        Case Else:     run = run & "\sa120\fs22 "
    End Select
    RTF_Parrafo = run & EscRTF(mk & " " & txt) & "\b0\par" & vbCrLf
End Function

Public Function EscRTF(ByVal s As String) As String
    s = Replace$(s, "\", "\\")
    s = Replace$(s, "{", "\{")
    s = Replace$(s, "}", "\}")
    s = Replace$(s, vbCrLf, "\par ")
    s = Replace$(s, vbLf, "\par ")
    s = Replace$(s, vbCr, "\par ")
    EscRTF = s
End Function

Private Sub SaveTextFile(ByVal fullPath As String, ByVal content As String)
    Dim ff As Integer: ff = FreeFile
    Open fullPath For Output As #ff
    Print #ff, content
    Close #ff
End Sub

Private Function DesktopPath() As String
    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & Application.PathSeparator
End Function

Private Function RutaAbs(ByVal basePath As String, ByVal ruta As String) As String
    If ruta Like "[A-Za-z]:\*" Or Left$(ruta, 2) = "\\" Then
        RutaAbs = ruta
    ElseIf Left$(ruta, 1) = "\" Or Left$(ruta, 1) = "/" Then
        RutaAbs = basePath & ruta
    Else
        RutaAbs = basePath & "\" & ruta
    End If
End Function

Private Function NzSafe(ByVal v As Variant) As String
    If IsError(v) Then NzSafe = "" Else NzSafe = Trim$(CStr(v))
End Function

Private Function LimpiarNombreArchivo(ByVal nombre As String) As String
    Dim ch As Variant, noP: noP = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    For Each ch In noP: nombre = Replace(nombre, ch, "_"): Next
    LimpiarNombreArchivo = nombre
End Function

Private Function EvitarColisiones(ByVal rutaCarpeta As String, ByVal nombreArchivo As String) As String
    Dim base$, ext$, cand$, n&
    SplitNameExt nombreArchivo, base, ext
    cand = base & ext
    If Dir$(rutaCarpeta & cand, vbNormal Or vbHidden Or vbSystem) = "" Then EvitarColisiones = cand: Exit Function
    n = 1
    Do
        cand = base & "_" & n & ext
        If Dir$(rutaCarpeta & cand, vbNormal Or vbHidden Or vbSystem) = "" Then EvitarColisiones = cand: Exit Function
        n = n + 1
    Loop
End Function

Private Sub SplitNameExt(ByVal nombre As String, ByRef base As String, ByRef ext As String)
    Dim p&: p = InStrRev(nombre, ".")
    If p > 0 Then base = Left$(nombre, p - 1): ext = Mid$(nombre, p) Else base = nombre: ext = ".rtf"
End Sub

' ======= validación mínima interfaz =======
Private Function ValidarInterfaz() As Boolean
    On Error GoTo ErrH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Interfaz")
    Dim e$, c$, p$: e = Trim$(CStr(ws.Range("E56").Value))
    c = Trim$(CStr(ws.Range("E57").Value))
    p = Trim$(CStr(ws.Range("E58").Value))
    If e = "" Or c = "" Or p = "" Then
        MsgBox "Faltan datos esenciales en Interfaz (E56:E58).", vbExclamation
        ValidarInterfaz = False: Exit Function
    End If
    ValidarInterfaz = True: Exit Function
ErrH:
    MsgBox "Error en ValidarInterfaz: " & Err.Description, vbCritical
    ValidarInterfaz = False
End Function


' ===========================
'  Gestión de Tablas (RSCIEI)
' ===========================

Private Function HojaNombres() As Variant
    HojaNombres = Array( _
        "Anx. II Sup", "Anx. II Reac", "Anx. II Res Sec", "Anx. II Res Ext", _
        "Anx. II Ocu", "Anx. II Sal", "Anx. II Res", _
        "Anx. III Det Auto", "Anx. III Det Man", "Anx. III Alr", _
        "Anx. III Meg", "Anx. III Hid", "Anx. III Ext Ext", "Anx. III BIES", _
        "Anx. III Ext Auto", "Anx. III Col", "Anx. III Hum")
End Function

Public Function EsTablaExcel(ByVal etiqueta As String) As Boolean
    Dim lo As ListObject
    Set lo = FindListObjectByEtiqueta(etiqueta, GetLibroCuna())
    EsTablaExcel = Not lo Is Nothing
End Function

Public Function ObtenerTablaPorEtiqueta(ByVal etiqueta As String) As Range
    Dim wb As Workbook
    Dim lo As ListObject
    
    Set wb = GetLibroCuna(True)
    Set lo = FindListObjectByEtiqueta(etiqueta, wb)
    
    If lo Is Nothing Then
        Err.Raise vbObjectError + 7401, "ObtenerTablaPorEtiqueta", _
                  "No se encontró ninguna tabla para la etiqueta: " & etiqueta
    End If
    
    Set ObtenerTablaPorEtiqueta = lo.Range
End Function

Public Sub PegarEtiquetaComoTablaEnWord(ByVal etiqueta As String, ByVal wdSelection As Object)
    Dim rng As Range
    Set rng = ObtenerTablaPorEtiqueta(etiqueta)
    
    rng.Copy
    wdSelection.PasteExcelTable LinkedToExcel:=False, WordFormatting:=True, rtf:=True
End Sub

Private Function FindListObjectByEtiqueta(ByVal etiqueta As String, ByVal wb As Workbook) As ListObject
    Dim lo As ListObject, ws As Worksheet
    Dim candidato As ListObject
    Dim norm As String, candidatos() As String, i As Long
    
    norm = NormalizarEtiqueta(etiqueta)
    candidatos = CandidateNames(norm)
    
    ' 1) Exacto por Name o DisplayName
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, etiqueta, vbTextCompare) = 0 _
               Or StrComp(lo.DisplayName, etiqueta, vbTextCompare) = 0 Then
                Set FindListObjectByEtiqueta = lo
                Exit Function
            End If
        Next lo
    Next ws
    
    ' 2) Por normalizado + prefijos típicos
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            For i = LBound(candidatos) To UBound(candidatos)
                If StrComp(lo.Name, candidatos(i), vbTextCompare) = 0 _
                   Or StrComp(lo.DisplayName, candidatos(i), vbTextCompare) = 0 Then
                    Set FindListObjectByEtiqueta = lo
                    Exit Function
                End If
            Next i
        Next lo
    Next ws
    
    ' 3) Coincidencia parcial (normalizado)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If InStr(1, NormalizarEtiqueta(lo.Name), norm, vbTextCompare) > 0 _
               Or InStr(1, NormalizarEtiqueta(lo.DisplayName), norm, vbTextCompare) > 0 Then
                If candidato Is Nothing Then Set candidato = lo
            End If
        Next lo
    Next ws
    
    If Not candidato Is Nothing Then
        Set FindListObjectByEtiqueta = candidato
    Else
        Set FindListObjectByEtiqueta = Nothing
    End If
End Function

Private Function GetLibroCuna(Optional ByVal forceOpenError As Boolean = False) As Workbook
    Dim wb As Workbook
    
    On Error Resume Next
    Set wb = Workbooks("EsquemaCuna.xlsx")
    On Error GoTo 0
    
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If wb Is Nothing And forceOpenError Then
        Err.Raise vbObjectError + 7402, "GetLibroCuna", _
                  "No se ha podido localizar el libro cuna (p.ej. EsquemaCuna.xlsx). Ábrelo o ajusta GetLibroCuna."
    End If
    
    Set GetLibroCuna = wb
End Function

Private Function NormalizarEtiqueta(ByVal s As String) As String
    Dim t As String
    t = s
    t = Trim$(t)
    
    t = ReplaceMulti(t, Array("á", "é", "í", "ó", "ú", "Á", "É", "Í", "Ó", "Ú", "ñ", "Ñ"), _
                        Array("a", "e", "i", "o", "u", "A", "E", "I", "O", "U", "n", "N"))
    t = ReplaceNonAlnumWithUnderscore(t)
    
    Do While InStr(t, "__") > 0
        t = Replace(t, "__", "_")
    Loop
    
    If Left$(t, 1) = "_" Then t = Mid$(t, 2)
    If Right$(t, 1) = "_" Then t = Left$(t, Len(t) - 1)
    
    NormalizarEtiqueta = t
End Function

Private Function CandidateNames(ByVal norm As String) As String()
    Dim arr() As String
    ReDim arr(0 To 5)
    arr(0) = norm
    arr(1) = "tbl_" & norm
    arr(2) = "t_" & norm
    arr(3) = Replace(norm, "_", "")
    arr(4) = "tbl" & Replace(norm, "_", "")
    arr(5) = "t" & Replace(norm, "_", "")
    CandidateNames = arr
End Function

Private Function ReplaceMulti(ByVal text As String, ByVal findArr As Variant, ByVal repArr As Variant) As String
    Dim i As Long, tmp As String
    tmp = text
    For i = LBound(findArr) To UBound(findArr)
        tmp = Replace(tmp, CStr(findArr(i)), CStr(repArr(i)))
    Next i
    ReplaceMulti = tmp
End Function

Private Function ReplaceNonAlnumWithUnderscore(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch Like "[A-Za-z0-9_]") Then
            out = out & ch
        ElseIf ch = " " Or ch = "-" Or ch = "." Or ch = "/" Or ch = "\" Or ch = ":" Then
            out = out & "_"
        Else
            out = out & "_"
        End If
    Next i
    ReplaceNonAlnumWithUnderscore = out
End Function

Private Function EtiquetaEsTablaCandidata(ByVal etiqueta As String) As Boolean
    Dim nombres As Variant, i As Long
    nombres = Array("Anx. II Sup", "Anx. II Reac", "Anx. II Res Sec", "Anx. II Res Ext", _
                    "Anx. II Ocu", "Anx. II Sal", "Anx. II Res", _
                    "Anx. III Det Auto", "Anx. III Det Man", "Anx. III Alr", _
                    "Anx. III Meg", "Anx. III Hid", "Anx. III Ext Ext", "Anx. III BIES", _
                    "Anx. III Ext Auto", "Anx. III Col", "Anx. III Hum")
    For i = LBound(nombres) To UBound(nombres)
        If StrComp(etiqueta, nombres(i), vbTextCompare) = 0 Then
            EtiquetaEsTablaCandidata = True
            Exit Function
        End If
    Next
End Function

Public Sub SustituirMarcadoresTablas(ByVal wdDoc As Object)
    Dim wdApp As Object: Set wdApp = wdDoc.Application
    Dim rFind As Object, rClose As Object, rWhole As Object
    Dim ok As Boolean, etiqueta As String
    Const INI$ = "@@TABLA:"
    Const fin$ = "@@"

    Set rFind = wdDoc.content.Duplicate

    With rFind.Find
        .ClearFormatting
        .text = INI
        .MatchWildcards = False
        .Forward = True
        .Wrap = 0 ' wdFindStop
    End With

    Do
        ok = rFind.Find.Execute
        If Not ok Then Exit Do

        Set rClose = wdDoc.Range(Start:=rFind.End, End:=wdDoc.content.End)
        With rClose.Find
            .ClearFormatting
            .text = fin
            .MatchWildcards = False
            .Forward = True
            .Wrap = 0
        End With
        If Not rClose.Find.Execute Then Exit Do

        Set rWhole = wdDoc.Range(Start:=rFind.Start, End:=rClose.End)

        etiqueta = rWhole.text
        etiqueta = Mid$(etiqueta, Len(INI) + 1)
        etiqueta = Left$(etiqueta, Len(etiqueta) - Len(fin))
        etiqueta = Trim$(etiqueta)

        rWhole.text = ""
        rWhole.Select
        PegarTablaDesdeCuna etiqueta, wdApp.Selection

        rFind.SetRange Start:=wdApp.Selection.Range.End, End:=wdDoc.content.End
    Loop
End Sub

Private Function ExtraerEtiquetaDeMarcador(ByVal marcador As String) As String
    Dim t$: t = marcador
    t = Replace(t, "@@TABLA:", "", , , vbTextCompare)
    t = Replace(t, "@@", "", , , vbTextCompare)
    ExtraerEtiquetaDeMarcador = Trim$(t)
End Function

Public Sub PegarTablaDesdeCuna(ByVal etiqueta As String, ByVal wdSelection As Object)
    Dim wb As Workbook, rng As Range
    Dim tbl As Object
    
    On Error GoTo fallo
    
    ' 1) Libro Cuna o ThisWorkbook
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks("EsquemaCuna.xlsx")
    On Error GoTo 0
    If wb Is Nothing Then
        If Not ThisWorkbook Is Nothing Then
            Set wb = ThisWorkbook
        Else
            Set wb = ActiveWorkbook
        End If
    End If
    If wb Is Nothing Then GoTo no_tabla
    
    ' 2) Rango apretado por etiqueta (respeta filtros visibles)
    Set rng = ResolverRangoApretadoPorEtiqueta(etiqueta, wb)
    If rng Is Nothing Then GoTo no_tabla
    
    ' 3) Copiar y pegar como tabla (Word formatea)
    rng.Copy
    wdSelection.PasteExcelTable LinkedToExcel:=False, WordFormatting:=True, rtf:=True
    
    ' 4) Ajustes de maquetación en Word
    On Error Resume Next
    Set tbl = wdSelection.Tables(1)
    If tbl Is Nothing Then
        Set tbl = wdSelection.Document.Tables(wdSelection.Document.Tables.Count)
    End If
    
    If Not tbl Is Nothing Then
        tbl.Rows.Alignment = 1                ' center
        tbl.AllowAutoFit = True
        tbl.AutoFitBehavior 2                 ' wdAutoFitWindow
        tbl.PreferredWidthType = 2            ' percent
        tbl.PreferredWidth = 100
        tbl.Rows.LeftIndent = 0
        tbl.TopPadding = 2
        tbl.BottomPadding = 2
        If tbl.Rows.Count >= 1 Then
            tbl.Rows(1).HeadingFormat = True
            tbl.Rows(1).Range.bold = True
        End If
        tbl.Borders.Enable = True
        tbl.Borders.OutsideLineStyle = 1
        tbl.Borders.InsideLineStyle = 1

        wdSelection.SetRange Start:=tbl.Range.End, End:=tbl.Range.End
        wdSelection.TypeParagraph
    End If
    On Error GoTo 0
    Exit Sub

no_tabla:
    wdSelection.TypeText "[NO SE ENCONTRÓ TABLA PARA: " & etiqueta & "]"
    wdSelection.TypeParagraph
    Exit Sub

fallo:
    On Error Resume Next
    wdSelection.TypeText "[ERROR PEGANDO TABLA (" & etiqueta & "): " & Err.Number & " - " & Err.Description & "]"
    wdSelection.TypeParagraph
End Sub

Private Function NormalizarEtiquetaLocal(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, "á", "a"): t = Replace(t, "é", "e"): t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o"): t = Replace(t, "ú", "u")
    t = Replace(t, "Á", "A"): t = Replace(t, "É", "E"): t = Replace(t, "Í", "I")
    t = Replace(t, "Ó", "O"): t = Replace(t, "Ú", "U")
    t = Replace(t, "ñ", "n"): t = Replace(t, "Ñ", "N")
    t = Replace(t, " ", "_"): t = Replace(t, ".", "_"): t = Replace(t, "-", "_")
    t = Replace(t, "/", "_"): t = Replace(t, "\", "_"): t = Replace(t, ":", "_")
    Do While InStr(t, "__") > 0: t = Replace(t, "__", "_"): Loop
    If Left$(t, 1) = "_" Then t = Mid$(t, 2)
    If Right$(t, 1) = "_" Then t = Left$(t, Len(t) - 1)
    NormalizarEtiquetaLocal = t
End Function

Private Function ResolverListObjectPorEtiqueta(ByVal etiqueta As String) As ListObject
    Dim wb As Workbook: Set wb = GetLibroCuna(True)
    Dim ws As Worksheet, lo As ListObject
    Dim norm$: norm = NormalizarEtiqueta(etiqueta)

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, etiqueta, vbTextCompare) = 0 _
            Or StrComp(lo.DisplayName, etiqueta, vbTextCompare) = 0 Then
                Set ResolverListObjectPorEtiqueta = lo
                Exit Function
            End If
        Next lo
    Next ws

    Dim cand As Variant, i&
    cand = CandidateNames(norm)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            For i = LBound(cand) To UBound(cand)
                If StrComp(NormalizarEtiqueta(lo.Name), cand(i), vbTextCompare) = 0 _
                Or StrComp(NormalizarEtiqueta(lo.DisplayName), cand(i), vbTextCompare) = 0 Then
                    Set ResolverListObjectPorEtiqueta = lo
                    Exit Function
                End If
            Next i
        Next lo
    Next ws
End Function

Private Function TempPath() As String
    Dim p$
    p = Environ$("TEMP")
    If Right$(p, 1) <> "\" And Right$(p, 1) <> "/" Then p = p & Application.PathSeparator
    TempPath = p
End Function

Private Sub DeleteIfExists(ByVal fullPath As String)
    On Error Resume Next
    If LenB(Dir$(fullPath)) > 0 Then Kill fullPath
    On Error GoTo 0
End Sub

' ====== Estilos PXL ======
Private Function AsegurarEstiloP(ByVal doc As Object, ByVal nombre As String, _
                                 Optional ByVal baseName As String = "Normal", _
                                 Optional ByVal outlineLevel As Long = 10) As Object
    Dim st As Object
    On Error Resume Next
    Set st = doc.Styles(nombre)
    On Error GoTo 0

    If st Is Nothing Then
        Set st = doc.Styles.add(Name:=nombre, Type:=1) ' párrafo
    End If

    On Error Resume Next
    If LCase$(baseName) = LCase$(nombre) Or Len(Trim$(baseName)) = 0 Then
        baseName = "Normal"
    End If
    st.BaseStyle = baseName
    If Err.Number = 5325 Then
        Err.Clear
        st.BaseStyle = "Normal"
    End If
    On Error GoTo 0

    st.ParagraphFormat.outlineLevel = outlineLevel
    Set AsegurarEstiloP = st
End Function

Public Sub ConfigurarEstilosPXL(ByVal doc As Object)
    Dim st As Object

    Set st = AsegurarEstiloP(doc, "Title", "Normal", 1)
    With st.Font
        .Size = 26
        .bold = False
        .Color = RGB(50, 62, 79)
    End With

    Set st = AsegurarEstiloP(doc, "PXL_Título 1", "Heading 1", 1)
    With st.Font
        .Name = "Century Gothic"
        .Size = 18
        .bold = True
        .Color = RGB(68, 114, 196)
    End With
    With st.ParagraphFormat
        .SpaceBefore = 6: .spaceAfter = 6: .LineSpacingRule = 0
        .keepWithNext = True
    End With

    Set st = AsegurarEstiloP(doc, "PXL_Título 2", "Heading 2", 2)
    With st.Font
        .Name = "Century Gothic"
        .Size = 16
        .bold = False
        .Color = RGB(0, 98, 166)
    End With
    With st.ParagraphFormat
        .SpaceBefore = 6: .spaceAfter = 6: .LineSpacingRule = 0
        .keepWithNext = True
    End With

    Set st = AsegurarEstiloP(doc, "PXL_Título 3", "Heading 3", 3)
    With st.Font
        .Name = "Century Gothic"
        .Size = 14
        .bold = False
        .Color = RGB(0, 98, 166)
    End With
    With st.ParagraphFormat
        .SpaceBefore = 6: .spaceAfter = 6: .LineSpacingRule = 0
        .keepWithNext = True
    End With

    Set st = AsegurarEstiloP(doc, "PXL_Párrafo", "Normal", 10)
    With st.Font
        .Name = "Century Gothic"
        .bold = False
        .Size = 10
        .Color = RGB(0, 0, 0)
    End With
    
    Set st = AsegurarEstiloP(doc, "PXL_Rojo", "Normal", 10)
    With st.Font
        .Name = "Century Gothic"
        .bold = False
        .Size = 10
        .Color = RGB(255, 0, 0)
    End With
    
End Sub

Private Sub GetMaxSizeAndAnyBold(ByVal rng As Object, ByRef maxSize As Single, ByRef anyBold As Boolean)
    Dim w As Object
    maxSize = 0: anyBold = False
    For Each w In rng.Words
        If w.text <> vbCr And w.text <> vbLf Then
            If w.Font.Size > maxSize Then maxSize = w.Font.Size
            If w.Font.bold = True Then anyBold = True
        End If
    Next w
End Sub

Public Sub AplicarEstilosPXLSegunRTF(ByVal wdDoc As Object)
    Dim p As Object, ch As Object
    Dim maxSize As Single, anyBold As Boolean, t As String

    On Error Resume Next

    For Each p In wdDoc.Paragraphs
        t = Trim$(Replace(p.Range.text, Chr$(13), ""))
        If Len(t) = 0 Then GoTo siguiente

        maxSize = 0: anyBold = False
        For Each ch In p.Range.Characters
            If ch.text <> vbCr And ch.text <> vbLf Then
                If ch.Font.Size > maxSize Then maxSize = ch.Font.Size
                If ch.Font.bold = True Then anyBold = True
            End If
        Next ch

        Dim targetStyle As String: targetStyle = ""
        If anyBold Then
            If maxSize >= 17.5 Then
                targetStyle = "PXL_Título 1"
            ElseIf maxSize >= 15 Then
                targetStyle = "PXL_Título 2"
            ElseIf maxSize >= 12 Then
                targetStyle = "PXL_Título 3"
            End If
        End If

        If Len(targetStyle) > 0 Then
            p.Range.ClearFormatting
            p.Range.Style = targetStyle
            p.Range.Font.Reset
        End If

siguiente:
    Next p

    On Error GoTo 0
End Sub

Public Sub AplicarEstilosPorMarcadores(ByVal wdDoc As Object)
    Dim mk As Variant
    For Each mk In Array("[[H1]]", "[[H2]]", "[[H3]]", "[[P]]", "[[R]]") ' << añade [[R]]
        Call AplicarEstiloParaMarcador(wdDoc, CStr(mk))
    Next mk
End Sub

Private Sub AplicarEstiloParaMarcador(ByVal wdDoc As Object, ByVal mk As String)
    Dim rng As Object, para As Object, sty$, mr As Object
    Dim startPos As Long, oneChar As Object

    Select Case mk
        Case "[[H1]]": sty = "PXL_Título 1"
        Case "[[H2]]": sty = "PXL_Título 2"
        Case "[[H3]]": sty = "PXL_Título 3"
        Case "[[R]]":  sty = "PXL_Rojo"        ' << NUEVO: estilo rojo
        Case Else:     sty = "PXL_Párrafo"
    End Select

    Set rng = wdDoc.content.Duplicate
    With rng.Find
        .ClearFormatting
        .text = mk
        .MatchWildcards = False
        .Forward = True
        .Wrap = 0
    End With

    Do While rng.Find.Execute
        Set para = rng.Paragraphs(1).Range

        startPos = rng.Start

        Set mr = wdDoc.Range(Start:=startPos, End:=startPos + Len(mk))
        mr.text = ""

        Set oneChar = wdDoc.Range(Start:=startPos, End:=startPos + 1)
        If oneChar.text = " " Then oneChar.text = ""

        With para
            .Font.Reset
            .ParagraphFormat.Reset
            .Style = sty
            .Font.Reset
        End With

        rng.SetRange Start:=para.End, End:=wdDoc.content.End
    Loop
End Sub


' ===== OPCIONAL: Sustitución de tokens dentro de Word (tablas/headers/footers) =====
Public Sub SustituirTokensEnTablasDeWord(ByVal wdDoc As Object, ByRef tokens() As TToken)
    On Error Resume Next
    Dim t As Object, r As Object, c As Object
    Dim txt As String
    If (Not Not tokens) = 0 Then Exit Sub

    For Each t In wdDoc.Tables
        For Each r In t.Rows
            For Each c In r.Cells
                txt = c.Range.text
                txt = QuitarFinDeCelda(txt)
                txt = ResolverConRegexEnTexto(txt, tokens) ' En Word no escapamos a RTF
                c.Range.text = txt
            Next c
        Next r
    Next t
End Sub

Private Function QuitarFinDeCelda(ByVal s As String) As String
    Dim t As String: t = s
    If Len(t) >= 2 Then
        If Right$(t, 2) = Chr$(13) & Chr$(7) Then t = Left$(t, Len(t) - 2)
    End If
    QuitarFinDeCelda = t
End Function

Private Function ResolverConRegexEnTexto(ByVal s As String, ByRef tokens() As TToken) As String
    Dim i As Long, re As Object, repl As String, out As String
    out = s
    For i = LBound(tokens) To UBound(tokens)
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = True
        re.Pattern = tokens(i).Pattern
        repl = tokens(i).Replacement
        out = re.Replace(out, repl)
    Next i
    ResolverConRegexEnTexto = out
End Function

Public Sub SustituirTokensEnCuerpoYSecciones(ByVal wdDoc As Object, ByRef tokens() As TToken)
    On Error Resume Next
    Dim rng As Object, s As Object
    If (Not Not tokens) = 0 Then Exit Sub

    Set rng = wdDoc.content.Duplicate
    ReemplazarTokensEnRangoWord rng, tokens

    For Each s In wdDoc.Sections
        ReemplazarTokensEnRangoWord s.Headers(1).Range, tokens ' Primary
        ReemplazarTokensEnRangoWord s.Headers(2).Range, tokens ' FirstPage
        ReemplazarTokensEnRangoWord s.Headers(3).Range, tokens ' EvenPages
        ReemplazarTokensEnRangoWord s.Footers(1).Range, tokens
        ReemplazarTokensEnRangoWord s.Footers(2).Range, tokens
        ReemplazarTokensEnRangoWord s.Footers(3).Range, tokens
    Next s
End Sub

Private Sub ReemplazarTokensEnRangoWord(ByVal rng As Object, ByRef tokens() As TToken)
    Dim txt As String
    txt = rng.text
    txt = ResolverConRegexEnTexto(txt, tokens)
    rng.text = txt
End Sub

Public Sub AplicarFormatoMemoria(ByVal wdDoc As Object)
    On Error Resume Next
    
    Const wdAlignParagraphJustify As Long = 3
    Const wdLineSpace1pt5 As Long = 1
    
    ' A) Todo el contenido del documento
    With wdDoc.content.ParagraphFormat
        .Alignment = wdAlignParagraphJustify
        .LineSpacingRule = wdLineSpace1pt5
        .SpaceBefore = 0
        .spaceAfter = 0
    End With
    
    ' B) Tablas
    Dim tbl As Object
    For Each tbl In wdDoc.Tables
        With tbl.Range.ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .spaceAfter = 0
        End With
    Next tbl
    
    ' C) Todas las historias (encabezados, pies, cuadros de texto…)
    Dim rng As Object
    For Each rng In wdDoc.StoryRanges
        With rng.ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .spaceAfter = 0
        End With
        Do While Not rng.NextStoryRange Is Nothing
            Set rng = rng.NextStoryRange
            With rng.ParagraphFormat
                .Alignment = wdAlignParagraphJustify
                .LineSpacingRule = wdLineSpace1pt5
                .SpaceBefore = 0
                .spaceAfter = 0
            End With
        Loop
    Next rng
    
    ' D) Estilos de párrafo (para que futuros textos con estilos queden igual)
    FormatearTodosLosEstilosDeParrafo wdDoc
End Sub
Private Sub FormatearTodosLosEstilosDeParrafo(ByVal wdDoc As Object)
    Const wdAlignParagraphJustify As Long = 3
    Const wdLineSpace1pt5 As Long = 1
    
    Dim st As Object
    For Each st In wdDoc.Styles
        ' 1 = wdStyleTypeParagraph (evitamos la constante)
        If st.Type = 1 Then
            With st.ParagraphFormat
                .Alignment = wdAlignParagraphJustify
                .LineSpacingRule = wdLineSpace1pt5
                .SpaceBefore = 0
                .spaceAfter = 0
            End With
        End If
    Next st
End Sub
Private Function ReemplazarTokensIMG(ByVal s As String) As String
    ' Convierte {IMG:Clave} -> @@IMG:Clave@@  (sin tocar otros tokens)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\{IMG:([^\{\}]+)\}"
    re.Global = True
    ReemplazarTokensIMG = re.Replace(s, "@@IMG:$1@@")
End Function

Public Sub SustituirMarcadoresImagenes(ByVal wdDoc As Object)
    Dim wdApp As Object: Set wdApp = wdDoc.Application
    Dim rngStart As Object, rngEnd As Object, rngWhole As Object
    Dim etiqueta As String, ruta As String
    Const INI$ = "@@IMG:"
    Const fin$ = "@@"

    Set rngStart = wdDoc.content.Duplicate

    With rngStart.Find
        .ClearFormatting
        .text = INI
        .MatchWildcards = False
        .Forward = True
        .Wrap = 0 'wdFindStop
    End With

    Do While rngStart.Find.Execute
        ' Buscar el cierre @@ a partir del final del inicio encontrado
        Set rngEnd = wdDoc.Range(Start:=rngStart.End, End:=wdDoc.content.End)
        With rngEnd.Find
            .ClearFormatting
            .text = fin
            .MatchWildcards = False
            .Forward = True
            .Wrap = 0 'wdFindStop
        End With
        If Not rngEnd.Find.Execute Then Exit Do

        ' Rango completo del marcador @@IMG:...@@
        Set rngWhole = wdDoc.Range(Start:=rngStart.Start, End:=rngEnd.End)

        ' Extraer la etiqueta entre @@IMG: y @@
        etiqueta = rngWhole.text
        etiqueta = Mid$(etiqueta, Len(INI) + 1)                       ' quita @@IMG:
        etiqueta = Left$(etiqueta, Len(etiqueta) - Len(fin))          ' quita @@
        etiqueta = LimpiarRuta(Trim$(etiqueta))                       ' limpia basura

        ' Resolver ruta:
        ' - si la “etiqueta” ya parece una ruta (tiene \ o : o / y extensión) => úsala
        ' - si no, se asume Clave => buscar en hoja Imagenes (u orígenes que tengas)
        If EsRutaProbable(etiqueta) Then
            ruta = NormalizarRuta(etiqueta)
        Else
            ruta = ObtenerRutaImagenPorClave(etiqueta) ' <- tu resolver de clave->ruta
        End If

        ' Reemplazar el marcador por la imagen o por mensaje de error legible
        rngWhole.text = ""                ' quita el marcador
        rngWhole.Select

        If Len(ruta) > 0 And ArchivoExiste(ruta) Then
            On Error Resume Next
            wdApp.Selection.InlineShapes.AddPicture _
                fileName:=ruta, LinkToFile:=False, SaveWithDocument:=True
            ' Ajuste opcional de tamaño máx (ancho en puntos). Quita si no lo quieres.
            If wdApp.Selection.InlineShapes.Count > 0 Then
                wdApp.Selection.InlineShapes(1).LockAspectRatio = True
                If wdApp.Selection.InlineShapes(1).Width > 420 Then ' ~14,8 cm
                    wdApp.Selection.InlineShapes(1).Width = 420
                End If
            End If
            On Error GoTo 0
        Else
            wdApp.Selection.TypeText "[Imagen no encontrada: " & etiqueta & "]"
        End If

        ' Continuar búsqueda a partir del final de lo recién insertado
        rngStart.SetRange Start:=wdApp.Selection.Range.End, End:=wdDoc.content.End
        With rngStart.Find
            .text = INI
        End With
    Loop
End Sub

Private Function ObtenerRutaImagenPorClave(ByVal clave As String) As String
    Dim wb As Workbook, ws As Worksheet
    Dim hdrRow As Long, colClave As Long, colRuta As Long
    Dim rngClaves As Range, m As Variant, fila As Long
    Dim ruta As String

    clave = Trim$(CStr(clave))
    If Len(clave) = 0 Then ObtenerRutaImagenPorClave = "": Exit Function

    ' 0) Workbook de datos: preferimos ThisWorkbook; si es add-in, usa ActiveWorkbook
    If Not ThisWorkbook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = ActiveWorkbook
    End If

    ' 1) Nombre definido IMG_<Clave> (si lo usas)
    On Error Resume Next
    ruta = GetRutaDesdeCelda(ws.Cells(fila, colRuta))
    On Error GoTo 0
    If Len(Trim$(ruta)) > 0 Then
        ObtenerRutaImagenPorClave = NormalizarRuta(ruta)
        Exit Function
    End If

    ' 2) Hoja llamada exactamente "Imagenes" (si existe)
    On Error Resume Next
    Set ws = wb.Worksheets("Imagenes")
    On Error GoTo 0
    If Not ws Is Nothing Then
        If DetectarCabecerasClaveRuta(ws, hdrRow, colClave, colRuta) Then
            Set rngClaves = ws.Range(ws.Cells(hdrRow + 1, colClave), ws.Cells(ws.Rows.Count, colClave).End(xlUp))
            m = Application.Match(clave, rngClaves, 0)
            If IsError(m) Then m = BuscarLineaManual(rngClaves, clave)
            If Not IsError(m) Then
                fila = rngClaves.row + CLng(m) - 1
                ruta = CStr(ws.Cells(fila, colRuta).Value)
                If Len(Trim$(ruta)) > 0 Then
                    ObtenerRutaImagenPorClave = NormalizarRuta(ruta)
                    Exit Function
                End If
            End If
        End If
    End If

    ' 3) Búsqueda global: cualquier hoja con cabeceras Clave/Ruta
    For Each ws In wb.Worksheets
        If DetectarCabecerasClaveRuta(ws, hdrRow, colClave, colRuta) Then
            Set rngClaves = ws.Range(ws.Cells(hdrRow + 1, colClave), ws.Cells(ws.Rows.Count, colClave).End(xlUp))
            If rngClaves.Rows.Count > 0 Then
                m = Application.Match(clave, rngClaves, 0)
                If IsError(m) Then m = BuscarLineaManual(rngClaves, clave)
                If Not IsError(m) Then
                    fila = rngClaves.row + CLng(m) - 1
                    ruta = CStr(ws.Cells(fila, colRuta).Value)
                    If Len(Trim$(ruta)) > 0 Then
                        ObtenerRutaImagenPorClave = NormalizarRuta(ruta)
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ws

    ' 4) Nada encontrado
    ObtenerRutaImagenPorClave = ""
End Function
'--- Busca la tabla tImagenes en todo el libro y devuelve Ruta por Clave ---
Private Function BuscarRutaEnTablaImagenes(ByVal clave As String) As String
    Dim ws As Worksheet, lo As ListObject, r As ListRow
    BuscarRutaEnTablaImagenes = ""
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, "tImagenes", vbTextCompare) = 0 Then
                For Each r In lo.ListRows
                    If StrComp(CStr(r.Range.Columns(1).Value), clave, vbTextCompare) = 0 Then
                        BuscarRutaEnTablaImagenes = CStr(r.Range.Columns(2).Value)
                        Exit Function
                    End If
                Next r
            End If
        Next lo
    Next ws
    On Error GoTo 0
End Function

'--- Recorre hojas “planas”: detecta cabeceras Clave/Ruta y busca la fila de la clave ---
Private Function BuscarRutaEnHojasSimples(ByVal clave As String) As String
    Dim ws As Worksheet, hdrRow As Long, colClave As Long, colRuta As Long
    Dim rngClave As Range, m As Variant, ruta As String
    
    BuscarRutaEnHojasSimples = ""
    For Each ws In ThisWorkbook.Worksheets
        If DetectarCabecerasClaveRuta(ws, hdrRow, colClave, colRuta) Then
            ' Rango de claves desde la fila siguiente a la cabecera
            Set rngClave = ws.Range(ws.Cells(hdrRow + 1, colClave), ws.Cells(ws.Rows.Count, colClave).End(xlUp))
            If rngClave.Rows.Count >= 1 Then
                m = Application.Match(clave, rngClave, 0)
                If Not IsError(m) Then
                    ruta = CStr(ws.Cells(rngClave.row + CLng(m) - 1, colRuta).Value)
                    If Len(Trim$(ruta)) > 0 Then
                        BuscarRutaEnHojasSimples = ruta
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ws
End Function

'--- Localiza cabeceras "Clave" y "Ruta" (case-insensitive) en las primeras filas/columnas ---
Private Function DetectarCabecerasClaveRuta(ByVal ws As Worksheet, _
                                            ByRef hdrRow As Long, _
                                            ByRef colClave As Long, _
                                            ByRef colRuta As Long) As Boolean
    Dim ur As Range, r As Long, c As Long, v As String
    Dim maxRows As Long, maxCols As Long
    
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function
    
    ' Escaneamos un área razonable al inicio
    maxRows = Application.Min(ur.Rows.Count, 20)
    maxCols = Application.Min(ur.Columns.Count, 30)
    
    hdrRow = 0: colClave = 0: colRuta = 0
    For r = ur.row To ur.row + maxRows - 1
        colClave = 0: colRuta = 0
        For c = ur.Column To ur.Column + maxCols - 1
            v = LCase$(Trim$(CStr(ws.Cells(r, c).Value)))
            If v = "clave" And colClave = 0 Then colClave = c
            If v = "ruta" And colRuta = 0 Then colRuta = c
            If colClave > 0 And colRuta > 0 Then
                hdrRow = r
                DetectarCabecerasClaveRuta = True
                Exit Function
            End If
        Next c
    Next r
End Function

'--- Normaliza rutas relativas/absolutas ---
Private Function NormalizarRuta(ByVal s As String) As String
    s = LimpiarRuta(s)
    ' Si termina con barra o no tiene extensión, no es archivo -> lo dejamos tal cual
    If Right$(s, 1) = "\" Or InStrRev(s, ".") = 0 Then
        NormalizarRuta = s
        Exit Function
    End If
    ' Absoluta o UNC
    If s Like "[A-Za-z]:\*" Or Left$(s, 2) = "\\" Then
        NormalizarRuta = s
    Else
        NormalizarRuta = RutaAbs(ThisWorkbook.path, s) ' relativa al .xlsm
    End If
End Function

Public Sub Test_IMG_Resolver()
    Dim clave As String, ruta As String
    clave = "PlanoGeneral"   ' <-- cambia para probar otras

    ruta = ObtenerRutaImagenPorClave(clave)

    If Len(ruta) = 0 Then
        MsgBox "No se encontró ruta para la clave: " & clave, vbExclamation, "IMG Test"
    Else
        Dim existe As String
        If Len(Dir$(ruta, vbNormal)) > 0 Then
            existe = "SÍ"
        Else
            existe = "NO"
        End If
        MsgBox "Clave: " & clave & vbCrLf & _
               "Ruta devuelta: " & ruta & vbCrLf & _
               "¿Existe el archivo?: " & existe, vbInformation, "IMG Test"
    End If
End Sub

Private Function BuscarLineaManual(ByVal rng As Range, ByVal clave As String) As Variant
    Dim r As Range
    For Each r In rng
        If StrComp(Trim$(CStr(r.Value)), Trim$(clave), vbTextCompare) = 0 Then
            BuscarLineaManual = r.row - rng.row + 1
            Exit Function
        End If
    Next r
    BuscarLineaManual = CVErr(xlErrNA)
End Function
Private Function GetRutaDesdeCelda(ByVal c As Range) As String
    Dim s As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        s = c.Hyperlinks(1).Address ' si es hipervínculo, usa la Address real
        If Len(s) = 0 Then s = c.Hyperlinks(1).SubAddress ' por si es interno
    Else
        s = CStr(c.Value2)
    End If
    On Error GoTo 0
    GetRutaDesdeCelda = LimpiarRuta(s)
End Function
Private Function LimpiarRuta(ByVal s As String) As String
    ' quita comillas, NBSP, tabs y saltos de línea ocultos
    s = Replace$(s, Chr$(160), " ") ' NBSP
    s = Replace$(s, Chr$(9), "")
    s = Replace$(s, Chr$(13), "")
    s = Replace$(s, Chr$(10), "")
    s = Replace$(s, """", "")
    s = Trim$(s)
    LimpiarRuta = s
End Function
Private Function ArchivoExiste(ByVal fullPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ArchivoExiste = fso.FileExists(fullPath)
End Function
Private Function EsRutaProbable(ByVal s As String) As Boolean
    s = LimpiarRuta(s)
    Dim tieneSep As Boolean: tieneSep = (InStr(s, "\") > 0) Or (InStr(s, "/") > 0) Or (InStr(s, ":") > 0)
    Dim tieneExt As Boolean: tieneExt = (InStrRev(s, ".") > 0 And Len(s) - InStrRev(s, ".") <= 5)
    EsRutaProbable = (tieneSep And tieneExt)
End Function
'--- Devuelve el rango recortado al contenido real (sin filas/columnas en blanco alrededor)
Private Function TrimRange(ByVal rng As Range) As Range
    On Error GoTo fallo
    If rng Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = rng.Worksheet
    Dim firstCell As Range, lastCell As Range
    Dim firstRow&, lastRow&, firstCol&, lastCol&
    
    ' Buscamos primera y última celda con algo (texto, número, fórmula) dentro del rng
    Set firstCell = rng.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                             SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If firstCell Is Nothing Then Exit Function ' todo vacío
    
    Set lastCell = rng.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    firstRow = firstCell.row: lastRow = lastCell.row
    
    Set firstCell = rng.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                             SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    Set lastCell = rng.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    firstCol = firstCell.Column: lastCol = lastCell.Column
    
    Set TrimRange = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
    Exit Function
fallo:
    Set TrimRange = rng ' si algo falla, devolvemos el original
End Function

'--- Rango “apretado” de un ListObject (respeta filas visibles si hay filtro)
Private Function TightRangeFromListObject(ByVal lo As ListObject, _
                                          Optional ByVal visibleOnly As Boolean = True) As Range
    Dim base As Range, vis As Range
    On Error Resume Next
    If lo.ShowHeaders Then
        If lo.DataBodyRange Is Nothing Then
            Set base = lo.HeaderRowRange
        Else
            Set base = lo.HeaderRowRange.Resize(lo.HeaderRowRange.Rows.Count + lo.DataBodyRange.Rows.Count, _
                                                Application.Max(lo.HeaderRowRange.Columns.Count, lo.DataBodyRange.Columns.Count))
        End If
    Else
        If lo.DataBodyRange Is Nothing Then
            Set base = lo.Range
        Else
            Set base = lo.DataBodyRange
        End If
    End If
    On Error GoTo 0
    If base Is Nothing Then Set base = lo.Range
    
    If visibleOnly Then
        On Error Resume Next
        Set vis = base.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not vis Is Nothing Then Set base = vis
    End If
    
    Set TightRangeFromListObject = TrimRange(base)
End Function

'--- Para hojas sin ListObject: recorta UsedRange o un rango candidato
Private Function TightRangeFromSheet(ByVal ws As Worksheet) As Range
    Dim ur As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Function
    Set TightRangeFromSheet = TrimRange(ur)
End Function

Private Function ResolverRangoApretadoPorEtiqueta(ByVal etiqueta As String, ByVal wb As Workbook) As Range
    Dim ws As Worksheet, lo As ListObject, rng As Range
    Dim etiquetaNorm As String, nombreNorm As String
    Dim encontrado As Boolean
    
    etiquetaNorm = NormalizarEtiquetaLocal(etiqueta)

    ' A) Coincidencia exacta con tablas
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, etiqueta, vbTextCompare) = 0 _
            Or StrComp(lo.DisplayName, etiqueta, vbTextCompare) = 0 Then
                Set ResolverRangoApretadoPorEtiqueta = TightRangeFromListObject(lo, True)
                Exit Function
            End If
        Next lo
    Next ws

    ' B) Normalizado + prefijos
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            nombreNorm = NormalizarEtiquetaLocal(lo.Name)
            If nombreNorm = etiquetaNorm _
               Or nombreNorm = "tbl_" & etiquetaNorm _
               Or nombreNorm = "t_" & etiquetaNorm _
               Or nombreNorm = Replace(etiquetaNorm, "_", "") _
               Or nombreNorm = "tbl" & Replace(etiquetaNorm, "_", "") _
               Or nombreNorm = "t" & Replace(etiquetaNorm, "_", "") Then
                Set ResolverRangoApretadoPorEtiqueta = TightRangeFromListObject(lo, True)
                Exit Function
            End If
        Next lo
    Next ws

    ' C) Hoja con nombre coincidente (o normalizado)
    On Error Resume Next
    Set ws = wb.Worksheets(etiqueta)
    On Error GoTo 0
    If ws Is Nothing Then
        For Each ws In wb.Worksheets
            If NormalizarEtiquetaLocal(ws.Name) = etiquetaNorm Then Exit For
        Next ws
    End If
    If Not ws Is Nothing Then
        ' Si hay tabla, usamos su rango apretado; si no, UsedRange apretado
        If ws.ListObjects.Count > 0 Then
            Set ResolverRangoApretadoPorEtiqueta = TightRangeFromListObject(ws.ListObjects(1), True)
        Else
            Set ResolverRangoApretadoPorEtiqueta = TightRangeFromSheet(ws)
        End If
        Exit Function
    End If
    
    ' D) Nada
    Set ResolverRangoApretadoPorEtiqueta = Nothing
End Function

