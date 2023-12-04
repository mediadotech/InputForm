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

'屋号
Private Const C0000_MSG As String = "なし"
Private Const C0001_MSG As String = "株式会社～"
Private Const C0002_MSG As String = "～株式会社"
Private Const C0003_MSG As String = "有限会社～"
Private Const C0004_MSG As String = "～有限会社"
Private Const C0005_MSG As String = "合同会社～"
Private Const C0006_MSG As String = "財団法人～"
Private Const C0007_MSG As String = "一般社団法人～"
Private Const C0008_MSG As String = "公益財団法人～"
Private Const C0009_MSG As String = "社団法人～"
Private Const C0010_MSG As String = "学校法人～"
Private Const C0011_MSG As String = "公益社団法人～"
Private Const C0012_MSG As String = "一般財団法人～"
Private Const C0013_MSG As String = "～合資会社"
Private Const C0014_MSG As String = "～合同会社"
Private Const C0015_MSG As String = "合資会社～"
Private Const C0016_MSG As String = "特定非営利活動法人～"
Private Const C0017_MSG As String = "国立大学法人～"

'発行形態
Private Const C0018_MSG As String = "コミック"
Private Const C0019_MSG As String = "書籍"
Private Const C0020_MSG As String = "写真集"
Private Const C0021_MSG As String = "雑誌"
Private Const C0022_MSG As String = "タテヨミ"

'取次会社名
Private Const PIF000_MSG As String = "出版デジタル機構"
Private Const PIF001_MSG As String = "ブックリスタ"
Private Const PIF002_MSG As String = "MBJ"
Private Const PIF003_MSG As String = "メディアドゥ"
Private Const PIF004_MSG As String = "電通"
Private Const PIF005_MSG As String = "富士山マガジンサービス"
Private Const PIF006_MSG As String = "角川コンテンツゲート"
Private Const PIF007_MSG As String = "md-dc"
Private Const PIF008_MSG As String = "緊デジ"

'入稿グループ
Private Const PIF009_MSG As String = "角川グループ"


Private Sub Class_Initialize()
    
    Set businessNameList = CreateObject("Scripting.Dictionary")
    Set publishTypeList = CreateObject("Scripting.Dictionary")
    Set agencyCompanyNameList = CreateObject("Scripting.Dictionary")
    Set summitGroupNameList = CreateObject("Scripting.Dictionary")
   
    '屋号
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
    
    '発行形態
    publishTypeList.Add C0018_MSG, C0018_MSG
    publishTypeList.Add C0019_MSG, C0018_MSG
    publishTypeList.Add C0020_MSG, C0018_MSG
    publishTypeList.Add C0021_MSG, C0018_MSG
    publishTypeList.Add C0022_MSG, C0018_MSG
    
    '取次会社名
    agencyCompanyNameList.Add PIF000_MSG, PIF000_MSG
    agencyCompanyNameList.Add PIF001_MSG, PIF001_MSG
    agencyCompanyNameList.Add PIF002_MSG, PIF002_MSG
    agencyCompanyNameList.Add PIF003_MSG, PIF003_MSG
    agencyCompanyNameList.Add PIF004_MSG, PIF004_MSG
    agencyCompanyNameList.Add PIF005_MSG, PIF005_MSG
    agencyCompanyNameList.Add PIF006_MSG, PIF006_MSG
    agencyCompanyNameList.Add PIF007_MSG, PIF007_MSG
    agencyCompanyNameList.Add PIF008_MSG, PIF008_MSG
    
    '入稿グループ
    summitGroupNameList.Add PIF009_MSG, PIF009_MSG
     
End Sub

'屋号のリストを取得
Public Property Get getBusinessNameList() As Object

    Set getBusinessNameList = businessNameList
        
End Property

'発行形態を取得
Public Property Get getPublishTypeList() As Object

    Set getPublishTypeList = publishTypeList

End Property


'取次会社名を取得
Public Property Get getAgencyCompanyNameList() As Object

    Set getAgencyCompanyNameList = agencyCompanyNameList
        
End Property
        

'入稿グループを取得
Public Property Get getSummitGroupNameList() As Object

       Set getSummitGroupNameList = summitGroupNameList
        
End Property

Private Sub Class_Terminate()
    Set businessNameList = Nothing
    Set publishTypeList = Nothing
    Set agencyCompanyNameList = Nothing
    Set summitGroupNameList = Nothing
End Sub

