VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "commonMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim errorStr As CommonErrorConst

Private Sub Class_Initialize()
    Set errorStr = New CommonErrorConst
End Sub

'必須入力警告
Public Sub showReqMsg(row As Integer, col As Integer)
    msgBox errorStr.ET000 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col, vbCritical
End Sub

'警告メッセージ
Public Sub showCommonWarnMsg(errorCode As String, row As Integer, col As Integer)
    
   Dim errorMsg As String
    
    Select Case errorCode
        '全角カタカナではない場合
        Case "ET001"
            errorMsg = errorStr.ET001 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        'ハイフン（-）込みの半角数字ではない場合
        Case "ET002"
            errorMsg = errorStr.ET002 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        '住所区分が正しくない場合
        Case "ET003"
            errorMsg = errorStr.ET003 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        '郵便番号が正しくない場合
        Case "ET004"
            errorMsg = errorStr.ET004 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        'メールアドレスが正しくない場合
        Case "ET005"
            errorMsg = errorStr.ET005 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        '発行形態が正しくない場合
        Case "ET006"
            errorMsg = errorStr.ET006 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        '屋号が正しくない場合
        Case "ET007"
            errorMsg = errorStr.ET007 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        '書店会社IDが正しくない
        Case "ET008"
            errorMsg = errorStr.ET008 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
         'サービス開始日が正しくない
        Case "ET009"
            errorMsg = errorStr.ET009 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
         'サービス終了日が正しくない
        Case "ET010"
            errorMsg = errorStr.ET010 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        Case "ET011"
            errorMsg = errorStr.ET011 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        '書店ID
        Case "ET012"
            errorMsg = errorStr.ET012 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
    End Select
    
    msgBox errorMsg, vbExclamation
    
End Sub

Private Sub Class_Terminate()
    Set errorStr = Nothing
End Sub
