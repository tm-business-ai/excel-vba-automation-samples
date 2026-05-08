Attribute VB_Name = "MonthlySummary"
Option Explicit

Private Const DATA_SHEET_NAME As String = "売上データ"
Private Const MONTH_SHEET_NAME As String = "月別集計"
Private Const PRODUCT_SHEET_NAME As String = "商品別集計"
Private Const STAFF_SHEET_NAME As String = "担当者別集計"
Private Const FIRST_DATA_ROW As Long = 2

' CSVを選択し、売上データシートへ取り込みます。
Public Sub ImportSalesCsv()
    Dim filePath As Variant
    Dim wsData As Worksheet
    Dim queryTable As QueryTable

    On Error GoTo ErrorHandler

    filePath = Application.GetOpenFilename("CSVファイル (*.csv),*.csv")
    If filePath = False Then Exit Sub

    Set wsData = PrepareSheet(DATA_SHEET_NAME)
    wsData.Cells.Clear

    Set queryTable = wsData.QueryTables.Add(Connection:="TEXT;" & CStr(filePath), Destination:=wsData.Range("A1"))
    With queryTable
        .TextFilePlatform = 65001
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1)
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    MsgBox "CSVの取込が完了しました。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "CSV取込でエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' 売上データから月別、商品別、担当者別の集計表を作成します。
Public Sub CreateMonthlySummary()
    Dim wsData As Worksheet
    Dim lastRow As Long

    On Error GoTo ErrorHandler

    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    If lastRow < FIRST_DATA_ROW Then
        MsgBox "集計対象の売上データがありません。", vbExclamation
        Exit Sub
    End If

    If Not ValidateSalesData(wsData, lastRow) Then Exit Sub

    CreateSummaryByMonth wsData, lastRow
    CreateSummaryByKey wsData, lastRow, 4, PRODUCT_SHEET_NAME, "商品名"
    CreateSummaryByKey wsData, lastRow, 5, STAFF_SHEET_NAME, "担当者"

    MsgBox "集計表の作成が完了しました。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "集計処理でエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' 売上日から年月を作成し、月別に売上金額を集計します。
Private Sub CreateSummaryByMonth(ByVal wsData As Worksheet, ByVal lastRow As Long)
    Dim summary As Object
    Dim rowIndex As Long
    Dim summaryKey As String
    Dim amount As Currency
    Dim wsSummary As Worksheet

    Set summary = CreateObject("Scripting.Dictionary")

    For rowIndex = FIRST_DATA_ROW To lastRow
        summaryKey = Format(CDate(wsData.Cells(rowIndex, 1).Value), "yyyy/mm")
        amount = CCur(wsData.Cells(rowIndex, 8).Value)
        AddAmount summary, summaryKey, amount
    Next rowIndex

    Set wsSummary = PrepareSheet(MONTH_SHEET_NAME)
    OutputSummary wsSummary, "売上月", summary
End Sub

' 指定された列の値をキーにして売上金額を集計します。
Private Sub CreateSummaryByKey(ByVal wsData As Worksheet, ByVal lastRow As Long, ByVal keyColumn As Long, ByVal sheetName As String, ByVal headerName As String)
    Dim summary As Object
    Dim rowIndex As Long
    Dim summaryKey As String
    Dim amount As Currency
    Dim wsSummary As Worksheet

    Set summary = CreateObject("Scripting.Dictionary")

    For rowIndex = FIRST_DATA_ROW To lastRow
        summaryKey = CStr(wsData.Cells(rowIndex, keyColumn).Value)
        amount = CCur(wsData.Cells(rowIndex, 8).Value)
        AddAmount summary, summaryKey, amount
    Next rowIndex

    Set wsSummary = PrepareSheet(sheetName)
    OutputSummary wsSummary, headerName, summary
End Sub

' Dictionaryへ売上金額を加算します。
Private Sub AddAmount(ByVal summary As Object, ByVal summaryKey As String, ByVal amount As Currency)
    If summary.Exists(summaryKey) Then
        summary(summaryKey) = summary(summaryKey) + amount
    Else
        summary.Add summaryKey, amount
    End If
End Sub

' 集計結果をシートへ出力します。
Private Sub OutputSummary(ByVal ws As Worksheet, ByVal keyHeader As String, ByVal summary As Object)
    Dim rowIndex As Long
    Dim key As Variant

    ws.Cells.Clear
    ws.Range("A1").Value = keyHeader
    ws.Range("B1").Value = "売上金額"

    rowIndex = 2
    For Each key In summary.Keys
        ws.Cells(rowIndex, 1).Value = key
        ws.Cells(rowIndex, 2).Value = summary(key)
        rowIndex = rowIndex + 1
    Next key

    ws.Columns("A:B").AutoFit
End Sub

' 売上データの必須項目、日付、金額を確認します。
Private Function ValidateSalesData(ByVal ws As Worksheet, ByVal lastRow As Long) As Boolean
    Dim rowIndex As Long
    Dim message As String

    For rowIndex = FIRST_DATA_ROW To lastRow
        If Not IsDate(ws.Cells(rowIndex, 1).Value) Then
            message = message & rowIndex & "行目: 売上日の日付形式を確認してください。" & vbCrLf
        End If

        If Trim(ws.Cells(rowIndex, 4).Value) = "" Then
            message = message & rowIndex & "行目: 商品名が未入力です。" & vbCrLf
        End If

        If Trim(ws.Cells(rowIndex, 5).Value) = "" Then
            message = message & rowIndex & "行目: 担当者が未入力です。" & vbCrLf
        End If

        If Not IsNumeric(ws.Cells(rowIndex, 8).Value) Then
            message = message & rowIndex & "行目: 売上金額は数値で入力してください。" & vbCrLf
        End If
    Next rowIndex

    If Len(message) > 0 Then
        MsgBox message, vbExclamation
        ValidateSalesData = False
    Else
        ValidateSalesData = True
    End If
End Function

' シートがなければ作成し、あれば初期化して返します。
Private Function PrepareSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set PrepareSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If PrepareSheet Is Nothing Then
        Set PrepareSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        PrepareSheet.Name = sheetName
    Else
        PrepareSheet.Cells.Clear
    End If
End Function
