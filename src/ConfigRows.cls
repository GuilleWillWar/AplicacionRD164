VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mItems As Collection

Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

Public Sub add(ByVal tipo As String, ByVal texto As String, ByVal salto As String)
    Dim it As RowItem
    Set it = New RowItem
    it.tipo = Trim$(tipo)
    it.texto = texto
    it.salto = UCase$(Trim$(salto))
    mItems.add it
End Sub

Public Property Get Item(ByVal idx As Long) As RowItem
    Set Item = mItems(idx)
End Property

Public Property Get Count() As Long
    Count = mItems.Count
End Property
