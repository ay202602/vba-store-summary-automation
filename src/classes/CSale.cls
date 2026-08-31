VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mProduct As String
Private mQuantity As Long
Private mUnitPrice As Currency

Property Let Product(v As Variant)
    If Trim(v) = "" Then
        Err.Raise vbObjectError + 1, "CProduct.Product", "商品名が入力されていません。"
    ElseIf VarType(v) <> vbString Then
        Err.Raise vbObjectError + 2, "CProduct.Product", "商品名は文字を入力してください。"
    End If
    mProduct = v
End Property

Property Get Product() As Variant
    Product = mProduct
End Property

Property Let Quantity(v As Long)
    If Trim(v) = "" Then
        Err.Raise vbObjectError + 3, "CProduct.Quantity", "数量が入力されていません。"
    ElseIf VarType(v) <> vbLong Then
        Err.Raise vbObjectError + 4, "CProduct.Quantity", "数量は数字を入力してください。"
    ElseIf v <= 0 Then
        Err.Raise vbObjectError + 5, "CProduct.Quantity", "数量は1以上必要です。"
    End If
    mQuantity = v
End Property

Property Get Quantity() As Long
    Quantity = mQuantity
End Property

Property Let UnitPrice(v As Currency)
    If Trim(v) = "" Then
        Err.Raise vbObjectError + 6, "CProduct.UnitPrice", "単価が入力されていません。"
    ElseIf VarType(v) <> vbCurrency Then
        Err.Raise vbObjectError + 7, "CProduct.UnitPrice", "単価は数字を入力してください。"
    ElseIf v <= 0 Then
        Err.Raise vbObjectError + 8, "CProduct.UnitPrice", "単価は1以上必要です。"
    End If
    mUnitPrice = v
End Property

Property Get UnitPrice() As Currency
    UnitPrice = mUnitPrice
End Property

Property Get TotalPrice() As Currency
    TotalPrice = Me.UnitPrice * Me.Quantity
End Property

Public Sub Init(Product As String, Quantity As Long, UnitPrice As Currency)
    Me.Product = Product
    Me.Quantity = Quantity
    Me.UnitPrice = UnitPrice
End Sub




