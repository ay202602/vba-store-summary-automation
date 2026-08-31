VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStoreName As String
Private mSales As Collection

Private Sub Class_Initialize()
    Set mSales = New Collection
End Sub

Property Let StoreName(v As String)
    If Trim(v) = "" Then
        Err.Raise vbObjectError + 100, "CStore.StoreName", "集計対象の店舗が存在しません。"
    End If
    mStoreName = v
End Property

Property Get StoreName() As String
    StoreName = mStoreName
End Property

Public Sub AddSale(v As CSale)
    mSales.Add v
End Sub

Property Get Sales() As Collection
    Set Sales = mSales
End Property





