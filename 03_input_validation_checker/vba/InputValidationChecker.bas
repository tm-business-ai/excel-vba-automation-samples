Attribute VB_Name = "InputValidationChecker"
Option Explicit

Private Const DATA_SHEET_NAME As String = "入力データ"
Private Const ERROR_SHEET_NAME As String = "エラー一覧"
Private Const FIRST_DATA_ROW As Long = 2
Private Const ERROR_COLOR As Long = 13421823

' 入力データの未入力、形式不備、重複をチェックします。
Public Sub CheckInputData()
    Dim wsData As Worksheet
    Dim wsError As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim errorRow As Long
    Dim duplicateKeys As Object
    Dim controlNo As String

    On Error GoTo ErrorHandler

    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    Set wsError = PrepareErrorSheet()
    Set duplicateKeys = CreateObject("Scripting.Dictionary")

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    ClearCheckResult wsData, lastRow

    If lastRow < FIRST_DATA_ROW Then
        MsgBox "チェック対象データがありません。", vbExclamation
        Exit Sub
    End If

    errorRow = 2

    For rowIndex = FIRST_DATA_ROW To lastRow
        controlNo = Trim(wsData.Cells(rowIndex, 1).Value)

        If controlNo = "" Then
            AddError wsError, errorRow, rowIndex, 1, "管理番号", "未入力です。", wsData
        ElseIf duplicateKeys.Exists(controlNo) Then
            AddError wsError, errorRow, rowIndex, 1, "管理番号", "重複しています。", wsData
        Else
            duplicateKeys.Add controlNo, rowIndex
        End If

        If Not IsDate(wsData.Cells(rowIndex, 2).Value) Then
            AddError wsError, errorRow, rowIndex, 2, "受付日", "日付形式を確認してください。", wsData
        End If

        If Trim(wsData.Cells(rowIndex, 3).Value) = "" Then
            AddError wsError, errorRow, rowIndex, 3, "部署", "未入力です。", wsData
        End If

        If Trim(wsData.Cells(rowIndex, 4).Value) = "" Then
            AddError wsError, errorRow, rowIndex, 4, "担当者", "未入力です。", wsData
        End If

        If Not IsNumeric(wsData.Cells(rowIndex, 5).Value) Then
            AddError wsError, errorRow, rowIndex, 5, "金額", "数値で入力してください。", wsData
        End If

        If Not IsNumeric(wsData.Cells(rowIndex, 6).Value) Or Val(wsData.Cells(rowIndex, 6).Value) <= 0 Then
            AddError wsError, errorRow, rowIndex, 6, "数量", "1以上の数値で入力してください。", wsData
        End If

        If Trim(wsData.Cells(rowIndex, 7).Value) = "" Then
            AddError wsError, errorRow, rowIndex, 7, "ステータス", "未入力です。", wsData
        End If
    Next rowIndex

    wsError.Columns("A:E").AutoFit

    If errorRow = 2 Then
        MsgBox "チェックが完了しました。エラーはありません。", vbInformation
    Else
        MsgBox "チェックが完了しました。エラー一覧を確認してください。", vbExclamation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "入力チェックでエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' エラー一覧へ内容を追加し、該当セルに色を付けます。
Private Sub AddError(ByVal wsError As Worksheet, ByRef errorRow As Long, ByVal dataRow As Long, ByVal dataColumn As Long, ByVal itemName As String, ByVal message As String, ByVal wsData As Worksheet)
    wsError.Cells(errorRow, 1).Value = dataRow
    wsError.Cells(errorRow, 2).Value = dataColumn
    wsError.Cells(errorRow, 3).Value = itemName
    wsError.Cells(errorRow, 4).Value = wsData.Cells(dataRow, dataColumn).Value
    wsError.Cells(errorRow, 5).Value = message

    wsData.Cells(dataRow, dataColumn).Interior.Color = ERROR_COLOR
    errorRow = errorRow + 1
End Sub

' 前回チェック時の色付けをクリアします。
Private Sub ClearCheckResult(ByVal wsData As Worksheet, ByVal lastRow As Long)
    If lastRow >= FIRST_DATA_ROW Then
        wsData.Range(wsData.Cells(FIRST_DATA_ROW, 1), wsData.Cells(lastRow, 7)).Interior.Pattern = xlNone
    End If
End Sub

' エラー一覧シートを作成または初期化します。
Private Function PrepareErrorSheet() As Worksheet
    On Error Resume Next
    Set PrepareErrorSheet = ThisWorkbook.Worksheets(ERROR_SHEET_NAME)
    On Error GoTo 0

    If PrepareErrorSheet Is Nothing Then
        Set PrepareErrorSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        PrepareErrorSheet.Name = ERROR_SHEET_NAME
    Else
        PrepareErrorSheet.Cells.Clear
    End If

    PrepareErrorSheet.Range("A1").Value = "行番号"
    PrepareErrorSheet.Range("B1").Value = "列番号"
    PrepareErrorSheet.Range("C1").Value = "項目名"
    PrepareErrorSheet.Range("D1").Value = "入力値"
    PrepareErrorSheet.Range("E1").Value = "エラー内容"
End Function
