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

Dim errorStr As errorConst

Private Sub Class_Initialize()
    Set errorStr = New errorConst
End Sub

'警告メッセージ
Public Sub showCommonWarnMsg(errorCode As String, row As Integer, col As Integer)
    
   Dim errorMsg As String
    
    Select Case errorCode
        '全角カタカナではない場合
        Case "E0001"
            errorMsg = errorStr.E0001 & vbCrLf & "行：" & row & vbCrLf & "列：" & Cells(1, col).Value
        'ハイフン（-）込みの半角数字ではない場合
        Case "E0002"
            errorMsg = errorStr.E0002 & vbCrLf & "行：" & row & vbCrLf & "列：" & Cells(1, col).Value
        '住所区分が正しくない場合
        Case "E0004"
            errorMsg = errorStr.E0004 & vbCrLf & "行：" & row & vbCrLf & "列：" & Cells(1, col).Value
        '住所区分が正しくない場合
        Case "E0005"
            errorMsg = errorStr.E0005 & vbCrLf & "行：" & row & vbCrLf & "列：" & Cells(1, col).Value
        'メールアドレスが正しくない場合
        Case "E0006"
            errorMsg = errorStr.E0006 & vbCrLf & "行：" & row & vbCrLf & "列：" & Cells(1, col).Value
    End Select
    
    msgBox errorMsg, vbExclamation
    
End Sub

Private Sub Class_Terminate()
    Set errorStr = Nothing
End Sub
