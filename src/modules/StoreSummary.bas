Attribute VB_Name = "StoreSummary"
Option Explicit

Sub ProductSummary()
    Const SUMMARY_SHEET_NAME As String = "集計"
    Const START_ROW As Long = 2
    
    On Error GoTo Errhandler

    Dim company As CCompany
    Set company = New CCompany

    Dim ws As Worksheet

    Dim errList As Collection
    Set errList = New Collection

    Dim currentSheetName As String
    Dim currentCell As String

    Dim storeCount As Long
    storeCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> SUMMARY_SHEET_NAME Then
            currentSheetName = ws.name

            Dim store As CStore
            Set store = New CStore
            
            store.StoreName = ws.name
            storeCount = storeCount + 1

            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

            Dim data As Variant
            data = ws.Range("A1:C" & lastRow).value

            Dim i As Long
            For i = START_ROW To UBound(data, 1)
                Dim rowHasError As Boolean
                rowHasError = False

                Dim sale As New CSale
                Set sale = New CSale
                
                currentCell = "A" & i
                Dim rawProduct As Variant
                rawProduct = data(i, 1)

                ' ユーザー側向けに標準モジュールにてバリデーション（クラス内は開発者向け）
                If Trim(rawProduct) = "" Then
                    errList.Add "商品名が空です。" & vbCrLf & _
                    "該当シート：" & ws.name & vbCrLf & _
                    "該当セル：" & "A" & i
                    rowHasError = True
                ElseIf VarType(rawProduct) <> vbString Then
                    errList.Add "商品名は文字列を入れてください。" & vbCrLf & _
                    "該当シート：" & ws.name & vbCrLf & _
                    "該当セル：" & "A" & i
                    rowHasError = True
                End If

                currentCell = "B" & i
                Dim rawQuantity As Variant
                rawQuantity = data(i, 2)

                If Trim(rawQuantity) = "" Then
                    errList.Add "数量が空欄です。" & vbCrLf & _
                    "該当シート：" & ws.name & vbCrLf & _
                    "該当セル：" & "B" & i
                    rowHasError = True
                ElseIf Not IsNumeric(rawQuantity) Then
                    errList.Add "数量は数字を入力してください。" & vbCrLf & _
                    "該当シート：" & ws.name & vbCrLf & _
                    "該当セル：" & "B" & i
                    rowHasError = True
                ElseIf rawQuantity <= 0 Then
                    errList.Add "数量は1以上必要です。" & vbCrLf & _
                    "該当シート：" & ws.name & vbCrLf & _
                    "該当セル：" & "B" & i
                    rowHasError = True
                End If

                currentCell = "C" & i
                Dim rawUnitPrice As Variant
                rawUnitPrice = data(i, 3)

                If Trim(rawUnitPrice) = "" Then
                    errList.Add "単価が空欄です。" & vbCrLf & _
                    "該当シート：" & ws.name & vbCrLf & _
                    "該当セル：" & "C" & i
                    rowHasError = True
                ElseIf Not IsNumeric(rawUnitPrice) Then
                    errList.Add "単価は数字を入力してください。" & vbCrLf & _
                    "該当シート：" & ws.name & vbCrLf & _
                    "該当セル：" & "B" & i
                    rowHasError = True
                ElseIf rawUnitPrice <= 0 Then
                    errList.Add "単価は1以上必要です。" & vbCrLf & _
                    "該当シート：" & ws.name & vbCrLf & _
                    "該当セル：" & "B" & i
                    rowHasError = True
                End If

                If rowHasError = False Then
                    currentCell = "A" & i & "~C" & i
                    sale.Init _
                        Product:=CStr(rawProduct), _
                        Quantity:=CLng(rawQuantity), _
                        UnitPrice:=CCur(rawUnitPrice)

                    'Cstore内のCollectionに追加
                    store.AddSale sale
                End If
            Next i
            'CCompany内のCollection内にCstoreを追加
            company.AddStore store
        End If
    Next ws
    
    If storeCount = 0 Then
        msgbox "集計対象の店舗が存在しません。", vbCritical, "店舗エラー"
        Exit Sub
    End If

    If errList.Count > 0 Then
        Dim msg As String
        Dim itm As Variant
        For Each itm In errList
            msg = msg & itm & vbCrLf & vbCrLf
        Next itm
        msgbox msg, vbCritical, "入力エラー"
        Exit Sub
    End If
    
    Dim result As Object
    Set result = company.SummarizeSalesByProduct()

    Dim wsSummary As Worksheet
    Set wsSummary = ThisWorkbook.Worksheets(SUMMARY_SHEET_NAME)
    
    Dim wsData() As Variant
    ReDim wsData(START_ROW To (START_ROW + result.Count - 1), 1 To 2)
    
    wsSummary.Cells.ClearContents
    wsSummary.Range("A1").value = "商品名"
    wsSummary.Range("B1").value = "合計売上"

    Dim outRow As Long
    outRow = START_ROW

    Dim key As Variant
    For Each key In result.keys
        wsData(outRow, 1) = key
        wsData(outRow, 2) = result(key)
        outRow = outRow + 1
    Next key
    
    ' -1は開始行をデータ１件目としてカウントするため代入
    wsSummary.Range("A2:B" & (START_ROW + result.Count - 1)).value = wsData
    wsSummary.columns("B").AutoFit

    ' 数字の表示形式を会計に変更
    wsSummary.columns("B").NumberFormatLocal = "_-\* #,##0_-;-\* #,##0_-;_-\* ""-""_-;_-@_-"

    wsSummary.Range("A1:B" & (START_ROW + result.Count - 1)).Sort _
        Key1:=wsSummary.Range("A1"), _
        Order1:=xlAscending, _
        Header:=xlYes
    Exit Sub

Errhandler:
    Select Case Err.Number
        Case 9
            msgbox "「集計」ワークシートを作成してください。", vbCritical, "集計シートエラー"
        Case vbObjectError + 1 To vbObjectError + 8
            msgbox Err.Description, vbCritical, "入力エラー"
        Case Else
            Dim errMsg As String
            errMsg = "想定外のエラーが発生しました。" & vbCrLf & _
                    "エラー番号：" & Err.Number & vbCrLf & _
                    "内容：" & Err.Description & vbCrLf & _
                    "発生元：" & Err.Source & vbCrLf & _
                    "対象シート：" & currentSheetName & vbCrLf & _
                    "対象セル：" & currentCell
            msgbox errMsg, vbCritical, "想定外エラー"
    End Select
    Exit Sub
End Sub

