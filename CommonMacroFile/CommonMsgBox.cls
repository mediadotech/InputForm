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

'�K�{���͌x��
Public Sub showReqMsg(row As Integer, col As Integer)
    msgBox errorStr.E0000 & vbCrLf & "���ځF" & Cells(1, col).Value & vbCrLf & "�����͍s�F" & row, vbCritical
End Sub

'�x�����b�Z�[�W
Public Sub showCommonWarnMsg(errorCode As String, row As Integer, col As Integer)
    
   Dim errorMsg As String
    
    Select Case errorCode
        '�S�p�J�^�J�i�ł͂Ȃ��ꍇ
        Case "E0001"
            errorMsg = errorStr.E0001 & vbCrLf & "���ځF" & Cells(1, col).Value & vbCrLf & "�s�F" & row & vbCrLf & "��F" & col
        '�n�C�t���i-�j���݂̔��p�����ł͂Ȃ��ꍇ
        Case "E0002"
            errorMsg = errorStr.E0002 & vbCrLf & "���ځF" & Cells(1, col).Value & vbCrLf & "�s�F" & row & vbCrLf & "��F" & col
        '�Z���敪���������Ȃ��ꍇ
        Case "E0003"
            errorMsg = errorStr.E0003 & vbCrLf & "���ځF" & Cells(1, col).Value & vbCrLf & "�s�F" & row & vbCrLf & "��F" & col
        '�X�֔ԍ����������Ȃ��ꍇ
        Case "E0004"
            errorMsg = errorStr.E0004 & vbCrLf & "���ځF" & Cells(1, col).Value & vbCrLf & "�s�F" & row & vbCrLf & "��F" & col
        '���[���A�h���X���������Ȃ��ꍇ
        Case "E0005"
            errorMsg = errorStr.E0005 & vbCrLf & "���ځF" & Cells(1, col).Value & vbCrLf & "�s�F" & row & vbCrLf & "��F" & col
        '���s�`�Ԃ��������Ȃ��ꍇ
        Case "E0006"
            errorMsg = errorStr.E0006 & vbCrLf & "���ځF" & Cells(1, col).Value & vbCrLf & "�s�F" & row & vbCrLf & "��F" & col
          '�������������Ȃ��ꍇ
        Case "E0007"
            errorMsg = errorStr.E0007 & vbCrLf & "���ځF" & Cells(1, col).Value & vbCrLf & "�s�F" & row & vbCrLf & "��F" & col
            
    End Select
    
    msgBox errorMsg, vbExclamation
    
End Sub

Private Sub Class_Terminate()
    Set errorStr = Nothing
End Sub