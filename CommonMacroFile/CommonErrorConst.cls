VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonErrorConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'<����>
'�K�{����
Private Const E0000_MSG As String = "�K�{���ڂ���͂��Ă�������"
'�J�^�J�i����
Private Const E0001_MSG As String = "���͂�����������܂���" & vbCrLf & "�S�p�J�^�J�i"
'�d�b�ԍ�
Private Const E0002_MSG As String = "�ԍ�������������܂���" & vbCrLf & "�i�n�C�t���i-�j���݂̔��p�����j"
'�Z���敪
Private Const E0003_MSG As String = "�Z���敪������������܂���" & vbCrLf & "���p�����A1�i�����j�A2�i���O�j"
'�X�֔ԍ�
Private Const E0004_MSG As String = "�X�֔ԍ�������������܂���" & vbCrLf & "7�����A���p����"
'���[���A�h���X
Private Const E0005_MSG As String = "���[���A�h���X������������܂���" & vbCrLf & "���p�A@���܂܂�Ă郁�[���`��"
'���s�`��
Private Const E0006_MSG As String = "���s�`�Ԃ�����������܂���"
'����
Private Const E0007_MSG As String = "�������������Ȃ�܂���"

'<�x����o�^>
'�x����R�[�h
Private Const E0100_MSG As String = "�x���捀�ڂ܂������R�[�h�����͂���Ă܂���"
Private Const E0101_MSG As String = "�x����R�[�h��5�����ɂ��Ă�������"
Private Const E0102_MSG As String = "�x����R�[�h�͔��p�p�����ł�"
'�o�^�ԍ�
Private Const E0103_MSG As String = "�o�^�ԍ���14�����ɂ��Ă�������"
Private Const E0104_MSG As String = "�o�^�ԍ��͔��p�p�����ł�"
Private Const E0105_MSG As String = "�o�^�ԍ��`�����Ⴂ�܂�" & vbCrLf & "���p�p��T�{���p����13��"
'�ǉ����̓t���O
Private Const E0106_MSG As String = "���s�`�Ԓǉ����̓t���O������������܂���" & vbCrLf & "���p�����A0�i�ǉ����̓t�Ȃ��j�A1�i�ǉ����͂���j"

'<�o�ŎГo�^>
'�x���於
Private Const E0200_MSG As String = "�o�ŎЍ��ڂ܂��x���於�����͂���Ă܂���"
'�o�ŎЌo���R�[�h�K�{�t���O
Private Const E0201_MSG As String = "�o�ŎЌo���R�[�h�K�{������������܂���" & vbCrLf & "���p�����A0�i�K�{�ł͂Ȃ��j�A1�i�K�{�j"
'�����t�@�C���Đ������t���O
Private Const E0202_MSG As String = "�����t�@�C���Đ������K�{������������܂���" & vbCrLf & "���p�����A0�i�Đ��������Ȃ��j�A1�i�Đ���������j"
'�掟��Ж�
Private Const E0203_MSG As String = "�掟��Ж����������Ȃ�܂���"
Private Const E0204_MSG As String = "��舵�����s�`�Ԃɓ��͂����`�Ԃ������͏o���܂�"
Private Const E0205_MSG As String = "��舵�����s�`�ԁA��̔��\���s�`�Ԃ����͂����`�Ԃ������͏o���܂�"
'���e�O���[�v
Private Const E0206_MSG As String = "���e�O���[�v���������Ȃ�܂���"
'�T�[�r�X�J�n��
Private Const E0207_MSG As String = "�T�[�r�X�J�n�����������Ȃ�܂���" & vbCrLf & "yyyy/mm/dd"
'�T�[�r�X�I����
Private Const E0208_MSG As String = "�T�[�r�X�I�������������Ȃ�܂���" & vbCrLf & "yyyy/mm/dd"
Private Const E0209_MSG As String = "�T�[�r�X�J�n�����O�̓��t�͓��͂ł��܂���"

'����
Public Property Get E0000() As String
    E0000 = E0000_MSG
End Property

Public Property Get E0001() As String
    E0001 = E0001_MSG
End Property

Public Property Get E0002() As String
    E0002 = E0002_MSG
End Property

Public Property Get E0003() As String
    E0003 = E0003_MSG
End Property

Public Property Get E0004() As String
    E0004 = E0004_MSG
End Property

Public Property Get E0005() As String
    E0005 = E0005_MSG
End Property

Public Property Get E0006() As String
    E0006 = E0006_MSG
End Property

Public Property Get E0007() As String
    E0007 = E0007_MSG
End Property


'�x����o�^
Public Property Get E0100() As String
    E0100 = E0100_MSG
End Property

Public Property Get E0101() As String
    E0101 = E0101_MSG
End Property

Public Property Get E0102() As String
    E0102 = E0102_MSG
End Property

Public Property Get E0103() As String
    E0103 = E0103_MSG
End Property

Public Property Get E0104() As String
    E0104 = E0104_MSG
End Property

Public Property Get E0105() As String
    E0105 = E0105_MSG
End Property

Public Property Get E0106() As String
    E0106 = E0106_MSG
End Property


'�o�ŎГo�^
Public Property Get E0200() As String
    E0200 = E0200_MSG
End Property

Public Property Get E0201() As String
    E0201 = E0201_MSG
End Property

Public Property Get E0202() As String
    E0202 = E0202_MSG
End Property

Public Property Get E0203() As String
    E0203 = E0203_MSG
End Property

Public Property Get E0204() As String
    E0204 = E0204_MSG
End Property

Public Property Get E0205() As String
    E0205 = E0205_MSG
End Property

Public Property Get E0206() As String
    E0206 = E0206_MSG
End Property

Public Property Get E0207() As String
    E0207 = E0207_MSG
End Property

Public Property Get E0208() As String
    E0208 = E0208_MSG
End Property

Public Property Get E0209() As String
    E0209 = E0209_MSG
End Property