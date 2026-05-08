Attribute VB_Name = "DataTransfer"
Option Explicit

Private Const INPUT_SHEET_NAME As String = "入力"
Private Const FORM_SHEET_NAME As String = "帳票"
Private Const FIRST_DATA_ROW As Long = 2

' 入力シートの内容を帳票シートへ転記します。
Public Sub TransferDataToForm()
    Dim wsInput As Worksheet
    Dim wsForm As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim outputRow As Long
    Dim errorMessage As String

    On Error GoTo ErrorHandler

    Set wsInput = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    Set wsForm = ThisWorkbook.Worksheets(FORM_SHEET_NAME)

    lastRow = GetLastRow(wsInput, 1)
    If lastRow < FIRST_DATA_ROW Then
        MsgBox "入力データがありません。", vbExclamation
        Exit Sub
    End If

    errorMessage = ValidateInputData(wsInput, lastRow)
    If Len(errorMessage) > 0 Then
        MsgBox errorMessage, vbExclamation
        Exit Sub
    End If

    ClearForm

    ' 帳票のヘッダー部分へ先頭データを転記します。
    wsForm.Range("B2").Value = wsInput.Cells(FIRST_DATA_ROW, 1).Value
    wsForm.Range("B3").Value = wsInput.Cells(FIRST_DATA_ROW, 2).Value
    wsForm.Range("B4").Value = wsInput.Cells(FIRST_DATA_ROW, 3).Value
    wsForm.Range("B5").Value = wsInput.Cells(FIRST_DATA_ROW, 4).Value

    outputRow = 9
    For rowIndex = FIRST_DATA_ROW To lastRow
        wsForm.Cells(outputRow, 1).Value = rowIndex - FIRST_DATA_ROW + 1
        wsForm.Cells(outputRow, 2).Value = wsInput.Cells(rowIndex, 5).Value
        wsForm.Cells(outputRow, 3).Value = wsInput.Cells(rowIndex, 6).Value
        wsForm.Cells(outputRow, 4).Value = wsInput.Cells(rowIndex, 7).Value
        wsForm.Cells(outputRow, 5).Value = wsInput.Cells(rowIndex, 8).Value
        outputRow = outputRow + 1
    Next rowIndex

    wsForm.Range("B6").Value = Now
    MsgBox "帳票への転記が完了しました。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "転記処理でエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' 帳票シートの出力範囲をクリアします。
Public Sub ClearForm()
    Dim wsForm As Worksheet

    Set wsForm = ThisWorkbook.Worksheets(FORM_SHEET_NAME)

    wsForm.Range("B2:B6").ClearContents
    wsForm.Range("A9:E100").ClearContents
End Sub

' 必須項目、数量、日付形式を確認します。
Private Function ValidateInputData(ByVal ws As Worksheet, ByVal lastRow As Long) As String
    Dim rowIndex As Long
    Dim message As String

    For rowIndex = FIRST_DATA_ROW To lastRow
        If Trim(ws.Cells(rowIndex, 1).Value) = "" Then
            message = message & rowIndex & "行目: 依頼番号が未入力です。" & vbCrLf
        End If

        If Not IsDate(ws.Cells(rowIndex, 2).Value) Then
            message = message & rowIndex & "行目: 依頼日の日付形式を確認してください。" & vbCrLf
        End If

        If Trim(ws.Cells(rowIndex, 3).Value) = "" Then
            message = message & rowIndex & "行目: 部署が未入力です。" & vbCrLf
        End If

        If Trim(ws.Cells(rowIndex, 4).Value) = "" Then
            message = message & rowIndex & "行目: 担当者が未入力です。" & vbCrLf
        End If

        If Trim(ws.Cells(rowIndex, 5).Value) = "" Then
            message = message & rowIndex & "行目: 品目が未入力です。" & vbCrLf
        End If

        If Not IsNumeric(ws.Cells(rowIndex, 6).Value) Or Val(ws.Cells(rowIndex, 6).Value) <= 0 Then
            message = message & rowIndex & "行目: 数量は1以上の数値で入力してください。" & vbCrLf
        End If

        If Not IsDate(ws.Cells(rowIndex, 7).Value) Then
            message = message & rowIndex & "行目: 希望納期の日付形式を確認してください。" & vbCrLf
        End If
    Next rowIndex

    ValidateInputData = message
End Function

' 指定列の最終行を取得します。
Private Function GetLastRow(ByVal ws As Worksheet, ByVal targetColumn As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, targetColumn).End(xlUp).Row
End Function
