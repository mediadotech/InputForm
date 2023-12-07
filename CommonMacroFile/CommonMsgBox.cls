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
    msgBox errorStr.E0000 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col, vbCritical
End Sub

'警告メッセージ
Public Sub showCommonWarnMsg(errorCode As String, row As Integer, col As Integer)
    
   Dim errorMsg As String
    
    Select Case errorCode
        '全角カタカナではない場合
        Case "E0001"
            errorMsg = errorStr.E0001 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        'ハイフン（-）込みの半角数字ではない場合
        Case "E0002"
            errorMsg = errorStr.E0002 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        '住所区分が正しくない場合
        Case "E0003"
            errorMsg = errorStr.E0003 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        '郵便番号が正しくない場合
        Case "E0004"
            errorMsg = errorStr.E0004 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        'メールアドレスが正しくない場合
        Case "E0005"
            errorMsg = errorStr.E0005 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
        '発行形態が正しくない場合
        Case "E0006"
            errorMsg = errorStr.E0006 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
          '屋号が正しくない場合
        Case "E0007"
            errorMsg = errorStr.E0007 & vbCrLf & "項目：" & Cells(1, col).Value & vbCrLf & "行：" & row & vbCrLf & "列：" & col
            
    End Select
    
    msgBox errorMsg, vbExclamation
    
End Sub

Private Sub Class_Terminate()
    Set errorStr = Nothing
End Sub
