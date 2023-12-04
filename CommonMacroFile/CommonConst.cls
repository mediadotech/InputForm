VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "commonConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim businessNameList As Object
Dim publishTypeList As Object
Dim agencyCompanyNameList As Object
Dim summitGroupNameList As Object

'����
Private Const C0000_MSG As String = "�Ȃ�"
Private Const C0001_MSG As String = "������Ё`"
Private Const C0002_MSG As String = "�`�������"
Private Const C0003_MSG As String = "�L����Ё`"
Private Const C0004_MSG As String = "�`�L�����"
Private Const C0005_MSG As String = "������Ё`"
Private Const C0006_MSG As String = "���c�@�l�`"
Private Const C0007_MSG As String = "��ʎВc�@�l�`"
Private Const C0008_MSG As String = "���v���c�@�l�`"
Private Const C0009_MSG As String = "�Вc�@�l�`"
Private Const C0010_MSG As String = "�w�Z�@�l�`"
Private Const C0011_MSG As String = "���v�Вc�@�l�`"
Private Const C0012_MSG As String = "��ʍ��c�@�l�`"
Private Const C0013_MSG As String = "�`�������"
Private Const C0014_MSG As String = "�`�������"
Private Const C0015_MSG As String = "������Ё`"
Private Const C0016_MSG As String = "�����c�������@�l�`"
Private Const C0017_MSG As String = "������w�@�l�`"

'���s�`��
Private Const C0018_MSG As String = "�R�~�b�N"
Private Const C0019_MSG As String = "����"
Private Const C0020_MSG As String = "�ʐ^�W"
Private Const C0021_MSG As String = "�G��"
Private Const C0022_MSG As String = "�^�e���~"

'�掟��Ж�
Private Const PIF000_MSG As String = "�o�Ńf�W�^���@�\"
Private Const PIF001_MSG As String = "�u�b�N���X�^"
Private Const PIF002_MSG As String = "MBJ"
Private Const PIF003_MSG As String = "���f�B�A�h�D"
Private Const PIF004_MSG As String = "�d��"
Private Const PIF005_MSG As String = "�x�m�R�}�K�W���T�[�r�X"
Private Const PIF006_MSG As String = "�p��R���e���c�Q�[�g"
Private Const PIF007_MSG As String = "md-dc"
Private Const PIF008_MSG As String = "�كf�W"

'���e�O���[�v
Private Const PIF009_MSG As String = "�p��O���[�v"


Private Sub Class_Initialize()
    
    Set businessNameList = CreateObject("Scripting.Dictionary")
    Set publishTypeList = CreateObject("Scripting.Dictionary")
    Set agencyCompanyNameList = CreateObject("Scripting.Dictionary")
    Set summitGroupNameList = CreateObject("Scripting.Dictionary")
   
    '����
    businessNameList.Add C0000_MSG, C0000_MSG
    businessNameList.Add C0001_MSG, C0001_MSG
    businessNameList.Add C0002_MSG, C0002_MSG
    businessNameList.Add C0003_MSG, C0003_MSG
    businessNameList.Add C0004_MSG, C0004_MSG
    businessNameList.Add C0005_MSG, C0005_MSG
    businessNameList.Add C0006_MSG, C0006_MSG
    businessNameList.Add C0007_MSG, C0007_MSG
    businessNameList.Add C0008_MSG, C0008_MSG
    businessNameList.Add C0009_MSG, C0009_MSG
    businessNameList.Add C0010_MSG, C0010_MSG
    businessNameList.Add C0011_MSG, C0011_MSG
    businessNameList.Add C0012_MSG, C0012_MSG
    businessNameList.Add C0013_MSG, C0013_MSG
    businessNameList.Add C0014_MSG, C0014_MSG
    businessNameList.Add C0015_MSG, C0015_MSG
    businessNameList.Add C0016_MSG, C0016_MSG
    businessNameList.Add C0017_MSG, C0017_MSG
    
    '���s�`��
    publishTypeList.Add C0018_MSG, C0018_MSG
    publishTypeList.Add C0019_MSG, C0018_MSG
    publishTypeList.Add C0020_MSG, C0018_MSG
    publishTypeList.Add C0021_MSG, C0018_MSG
    publishTypeList.Add C0022_MSG, C0018_MSG
    
    '�掟��Ж�
    agencyCompanyNameList.Add PIF000_MSG, PIF000_MSG
    agencyCompanyNameList.Add PIF001_MSG, PIF001_MSG
    agencyCompanyNameList.Add PIF002_MSG, PIF002_MSG
    agencyCompanyNameList.Add PIF003_MSG, PIF003_MSG
    agencyCompanyNameList.Add PIF004_MSG, PIF004_MSG
    agencyCompanyNameList.Add PIF005_MSG, PIF005_MSG
    agencyCompanyNameList.Add PIF006_MSG, PIF006_MSG
    agencyCompanyNameList.Add PIF007_MSG, PIF007_MSG
    agencyCompanyNameList.Add PIF008_MSG, PIF008_MSG
    
    '���e�O���[�v
    summitGroupNameList.Add PIF009_MSG, PIF009_MSG
     
End Sub

'�����̃��X�g���擾
Public Property Get getBusinessNameList() As Object

    Set getBusinessNameList = businessNameList
        
End Property

'���s�`�Ԃ��擾
Public Property Get getPublishTypeList() As Object

    Set getPublishTypeList = publishTypeList

End Property


'�掟��Ж����擾
Public Property Get getAgencyCompanyNameList() As Object

    Set getAgencyCompanyNameList = agencyCompanyNameList
        
End Property
        

'���e�O���[�v���擾
Public Property Get getSummitGroupNameList() As Object

       Set getSummitGroupNameList = summitGroupNameList
        
End Property

Private Sub Class_Terminate()
    Set businessNameList = Nothing
    Set publishTypeList = Nothing
    Set agencyCompanyNameList = Nothing
    Set summitGroupNameList = Nothing
End Sub
