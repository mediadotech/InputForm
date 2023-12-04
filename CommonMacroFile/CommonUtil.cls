VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "commonUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim errorStr As CommonErrorConst
Dim commonStr As commonConst

Private Sub Class_Initialize()
    Set errorStr = New CommonErrorConst
    Set commonStr = New commonConst
End Sub

'�S�p�m�F
Public Function hasZenkaku(s As String) As Boolean
  If Len(s) <> LenB(StrConv(s, vbFromUnicode)) Then
    hasZenkaku = True
  Else
    hasZenkaku = False
  End If
End Function

'������ɓ��{�ꂪ�܂܂�Ă��邩�m�F
Public Function findJapanese(str As String) As String
    '���K�\���N���X�I�u�W�F�N�g
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    '�������������{��ȊO�𒊏o
    reg.Pattern = "[^��-��@-����-ꞁV�X�Y�Z�-�]"
    '������̍Ō�܂Ō�������
    reg.Global = True
    '�w��Z���̓��{��ȊO���󕶎��ɒu��������
    findJapanese = reg.Replace(str, "")
      Set reg = Nothing
End Function

'�����񂪔��p�����ɂȂ��Ă��邩�m�F
Public Function checkNarrowNum(str As String) As String
    '���K�\���N���X�I�u�W�F�N�g
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    '�������������z����
    reg.Pattern = "[0-9]"
    '������̍Ō�܂Ō�������
    reg.Global = True
    '�w��Z���̔��z�������󕶎��ɒu��������
    checkNarrowNum = reg.Replace(str, "")
    Set reg = Nothing
End Function

'�����񂪃J�^�J�i�ɂȂ��Ă��邩�m�F
Public Function checkKatakana(str As String) As String
    '���K�\���N���X�I�u�W�F�N�g
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    '�����������J�^�J�i
    reg.Pattern = "[�@-��\�\\�[\�|]"
    '������̍Ō�܂Ō�������
    reg.Global = True
    '�w��Z���̃J�^�J�i���󕶎��ɒu��������
    checkKatakana = reg.Replace(str, "")
      Set reg = Nothing
End Function

'�����񂪔��p�����ƃn�C�t���ɂȂ��Ă��邩�m�F
Public Function checkNarrowNumAndHyphen(str As String) As String
    '���K�\���N���X�I�u�W�F�N�g
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    '�������������z����
    reg.Pattern = "[0-9\-]"
    '������̍Ō�܂Ō�������
    reg.Global = True
    '�w��Z���̔��z�������󕶎��ɒu��������
    checkNarrowNumAndHyphen = reg.Replace(str, "")
      Set reg = Nothing
End Function

'���K�\���Œl�����[���A�h���X�̌`�����m�F����
Private Function checkMailAdess(getVal As Variant) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    ' ���[���A�h���X�`���̐��K�\��
    reg.Pattern = "^[a-zA-Z0-9]+([a-zA-Z0-9\._-]){0,63}@[a-zA-Z0-9-]+\.(?:[\.a-zA-Z0-9-]+)*$"
   '������̍Ō�܂Ō�������
    reg.Global = True
   '�啶���Ə����������(True�F���Ȃ��AFlase�F����)
    reg.IgnoreCase = True
    checkMailAdess = reg.Test(getVal)
    Set reg = Nothing
End Function

'�����m�F
Public Function findBusinessName(str As String) As Boolean
    findBusinessName = commonStr.getBusinessNameList.Exists(str)
End Function

'���s�`�Ԋm�F
Public Function findPublishType(str As String) As Boolean
    findPublishType = commonStr.getPublishTypeList.Exists(str)
End Function

'�掟��Ж��m�F
Public Function findAgencyCompanyName(str As String) As Boolean
    findAgencyCompanyName = commonStr.getAgencyCompanyNameList.Exists(str)
End Function

'���e�O���[�v�m�F
Public Function findSummitGroupName(str As String) As Boolean
    findSummitGroupName = commonStr.getSummitGroupNameList.Exists(str)
End Function
    
'�S�p�J�^�J�i�m�F
Public Function checkWideKana(row As Integer, col As Integer) As String
    checkWideKana = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '���p���J�^�J�i�ł͂Ȃ�
    If Not hasZenkaku(targetStr) Or Len(checkKatakana(targetStr)) > 0 Then
        checkWideKana = "E0001"
        Exit Function
    End If
End Function

'�����m�F
Public Function checkBusinessName(row As Integer, col As Integer) As String
    checkBusinessName = "OK"
    Dim businessName As String
    businessName = Cells(row, col).Value
    If Not findBusinessName(businessName) Then
        checkBusinessName = "E0007"
        Exit Function
    End If
End Function

'�d�b�ԍ��m�F
Public Function checkNumberWithHyphen(row As Integer, col As Integer) As String
    checkNumberWithHyphen = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '�n�C�t�����܂܂�Ă�̂����鐔���ƃn�C�t���ȊO�̕���������
    If InStr(targetStr, "-") = 0 Or Len(checkNarrowNumAndHyphen(targetStr)) > 0 Then
        checkNumberWithHyphen = "E0002"
        Exit Function
    End If
End Function

'���[���A�h���X�m�F
Public Function checkMailAdressList(row As Integer, col As Integer) As String
    checkMailAdressList = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '�S�p�������܂܂�Ă���
    If hasZenkaku(targetStr) Then
        checkMailAdressList = "E0005"
        Exit Function
    End If
    If InStr(targetStr, ",") Then
        Dim mailAdressList As Variant
        Dim i As Long
        mailAdressList = Split(targetStr, ",")
        For i = LBound(mailAdressList) To UBound(mailAdressList)
            If Not checkMailAdess(mailAdressList(i)) Then
                checkMailAdressList = "E0005"
                Exit Function
            End If
        Next i
    Else
         If Not checkMailAdess(targetStr) Then
            checkMailAdressList = "E0005"
            Exit Function
        End If
    End If
End Function

'���s�`�ԍ��ڊm�F
Public Function checkPublishTypeList(row As Integer, col As Integer) As String
    checkPublishTypeList = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '�S�p�������܂܂�Ă���
    If InStr(targetStr, ",") Then
        Dim publishTypeList As Variant
        Dim i As Long
        publishTypeList = Split(targetStr, ",")
        For i = LBound(publishTypeList) To UBound(publishTypeList)
            Dim listStr As String
            listStr = publishTypeList(i)
            If Not findPublishType(listStr) Then
                checkPublishTypeList = "E0006"
                Exit Function
            End If
        Next i
    Else
         If Not findPublishType(targetStr) Then
            checkPublishTypeList = "E0006"
            Exit Function
        End If
    End If
    
End Function

'�Z���敪�m�F
Public Function checkAdressDivision(row As Integer, col As Integer) As String
    checkAdressDivision = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '���p�������P�܂��Q�ȊO�̐���������
    If Len(checkNarrowNum(targetStr)) > 0 Or Cells(row, col).Value < 1 Or Cells(row, col).Value > 2 Then
        checkAdressDivision = "E0003"
        Exit Function
    End If
End Function

'�X�֔ԍ��m�F
Public Function checkZipCode(row As Integer, col As Integer) As String
    checkZipCode = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '���p�����ł͂Ȃ�����7�����ł͂Ȃ�
    If Len(checkNarrowNum(targetStr)) > 0 Or Not Len(targetStr) = 7 Then
        checkZipCode = "E0004"
        Exit Function
    End If
End Function

Private Sub Class_Terminate()
    Set commonStr = Nothing
    Set errorStr = Nothing
End Sub