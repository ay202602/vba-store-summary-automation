VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStores As Collection

Private Sub Class_Initialize()
    Set mStores = New Collection
End Sub

Public Sub AddStore(v As CStore)
    mStores.Add v
End Sub

Public Function SummarizeSalesByProduct() As Object
    ' 論理エラー対策のため、バリデーションを追加
    If mStores.Count = 0 Then
        Err.Raise vbObjectError + 200, "CCompany.SummarizeSalesByProduct", "Store が1件も登録されていません。先にAddStoreを使用してください。"
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim store As CStore
    Dim sale As CSale
    
    ' 全CStoreを走査し、各CStoreが持つCsaleを商品（Product）ごとに集計
    For Each store In mStores
        For Each sale In store.Sales
            If dict.Exists(sale.Product) Then
                dict(sale.Product) = dict(sale.Product) + sale.TotalPrice
            Else
                dict.Add sale.Product, sale.TotalPrice
            End If
        Next sale
    Next store
    Set SummarizeSalesByProduct = dict
End Function
