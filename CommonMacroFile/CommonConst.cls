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
Dim serviceFileTypeList As Object
Dim proveStatusList As Object

'屋号(Trade name)
Private Const TN000_MSG As String = "なし"
Private Const TN001_MSG As String = "株式会社～"
Private Const TN002_MSG As String = "～株式会社"
Private Const TN003_MSG As String = "有限会社～"
Private Const TN004_MSG As String = "～有限会社"
Private Const TN005_MSG As String = "合同会社～"
Private Const TN006_MSG As String = "財団法人～"
Private Const TN007_MSG As String = "一般社団法人～"
Private Const TN008_MSG As String = "公益財団法人～"
Private Const TN009_MSG As String = "社団法人～"
Private Const TN010_MSG As String = "学校法人～"
Private Const TN011_MSG As String = "公益社団法人～"
Private Const TN012_MSG As String = "一般財団法人～"
Private Const TN013_MSG As String = "～合資会社"
Private Const TN014_MSG As String = "～合同会社"
Private Const TN015_MSG As String = "合資会社～"
Private Const TN016_MSG As String = "特定非営利活動法人～"
Private Const TN017_MSG As String = "国立大学法人～"

'発行形態(Publish Type)
Private Const PT001_MSG As String = "コミック"
Private Const PT002_MSG As String = "書籍"
Private Const PT003_MSG As String = "写真集"
Private Const PT004_MSG As String = "雑誌"
Private Const PT005_MSG As String = "タテヨミ"

'取次会社名(Agency company name)
Private Const ACN000_MSG As String = "出版デジタル機構"
Private Const ACN001_MSG As String = "ブックリスタ"
Private Const ACN002_MSG As String = "MBJ"
Private Const ACN003_MSG As String = "メディアドゥ"
Private Const ACN004_MSG As String = "電通"
Private Const ACN005_MSG As String = "富士山マガジンサービス"
Private Const ACN006_MSG As String = "角川コンテンツゲート"
Private Const ACN007_MSG As String = "md-dc"
Private Const ACN008_MSG As String = "緊デジ"

'入稿グループ(Summit Group)
Private Const SG001_MSG As String = "角川グループ"

'取り扱いファイル形式(Service File Type)
'IBOOKS（901001）
Private Const SFT001_MSG As String = "901001"
'XML（900901）
Private Const SFT002_MSG As String = "900901"
'MOBI（900801）
Private Const SFT003_MSG As String = "900801"
'TTX（900701）
Private Const SFT004_MSG As String = "900701"
'JPG(WEBTOON)（900602）
Private Const SFT005_MSG As String = "900602"
'JPG（900601）
Private Const SFT006_MSG As String = "900601"
'HDC（900501）
Private Const SFT007_MSG As String = "900501"
'TIFF（900401）
Private Const SFT008_MSG As String = "900401"
'PDF（900301）
Private Const SFT009_MSG As String = "900301"
'EPUB3.0(ブックライブ)（900217）
Private Const SFT010_MSG As String = "900217"
'EPUB3.0(ソルマーレ)（900216）
Private Const SFT011_MSG As String = "900216"
'EPUB3.0(講談社)（900215）
Private Const SFT012_MSG As String = "900215"
'EPUB3.0(LINE特典版)（900214）
Private Const SFT013_MSG As String = "900214"
'EPUB3.0(WEBTOON)（900213）
Private Const SFT014_MSG As String = "900213"
'EPUB3.0(BL)（900212）
Private Const SFT015_MSG As String = "900212"
'EPUB3.0(音声付)（900211）
Private Const SFT016_MSG As String = "900211"
'EPUB3.0(AP)（900210）
Private Const SFT017_MSG As String = "900210"
'EPUB3.0(出版社)（900209）
Private Const SFT018_MSG As String = "900209"
'EPUB3.0(YB)（900208）
Private Const SFT019_MSG As String = "900208"
'EPUB3.0(OMF)（900207）
Private Const SFT020_MSG As String = "900207"
'EPUB3.0(BW)（900206）
Private Const SFT021_MSG As String = "900206"
'EPUB3.0(2)集英（900205）
Private Const SFT022_MSG As String = "900205"
'EPUB3.0(1)集英（900204）
Private Const SFT023_MSG As String = "900204"
'EPUB3.0(追加なし)（900203）
Private Const SFT024_MSG As String = "900203"
'EPUB3.0(雑誌)（900202）
Private Const SFT025_MSG As String = "900202"
'EPUB2.0（900201）
Private Const SFT026_MSG As String = "900201"
'dotbook(講談社)（900103）
Private Const SFT027_MSG As String = "900103"
'dotbook（900102）
Private Const SFT028_MSG As String = "900102"
'dotbook（900101）
Private Const SFT029_MSG As String = "900101"
'XMDFneo（900004）
Private Const SFT030_MSG As String = "900004"
'XMDF3（900003）
Private Const SFT031_MSG As String = "900003"
'XMDF(HS)（900002）
Private Const SFT032_MSG As String = "900002"
'XMDF（900001）
Private Const SFT033_MSG As String = "900001"
'JPG（800001）
Private Const SFT034_MSG As String = "800001"
'プリントオンデマンド（700001）
Private Const SFT035_MSG As String = "700001"


'承認ステータス（Prove Status）
'未設定
Private Const PS001_MSG As String = "0"
'承認
Private Const PS002_MSG As String = "1"
'保留
Private Const PS003_MSG As String = "9"
'削除
Private Const PS004_MSG As String = "99"

Private Sub Class_Initialize()
    
    Set businessNameList = CreateObject("Scripting.Dictionary")
    Set publishTypeList = CreateObject("Scripting.Dictionary")
    Set agencyCompanyNameList = CreateObject("Scripting.Dictionary")
    Set summitGroupNameList = CreateObject("Scripting.Dictionary")
    Set serviceFileTypeList = CreateObject("Scripting.Dictionary")
    Set proveStatusList = CreateObject("Scripting.Dictionary")
   
    '屋号
    businessNameList.Add TN000_MSG, TN000_MSG
    businessNameList.Add TN001_MSG, TN001_MSG
    businessNameList.Add TN002_MSG, TN002_MSG
    businessNameList.Add TN003_MSG, TN003_MSG
    businessNameList.Add TN004_MSG, TN004_MSG
    businessNameList.Add TN005_MSG, TN005_MSG
    businessNameList.Add TN006_MSG, TN006_MSG
    businessNameList.Add TN007_MSG, TN007_MSG
    businessNameList.Add TN008_MSG, TN008_MSG
    businessNameList.Add TN009_MSG, TN009_MSG
    businessNameList.Add TN010_MSG, TN010_MSG
    businessNameList.Add TN011_MSG, TN011_MSG
    businessNameList.Add TN012_MSG, TN012_MSG
    businessNameList.Add TN013_MSG, TN013_MSG
    businessNameList.Add TN014_MSG, TN014_MSG
    businessNameList.Add TN015_MSG, TN015_MSG
    businessNameList.Add TN016_MSG, TN016_MSG
    businessNameList.Add TN017_MSG, TN017_MSG
    
    '発行形態
    publishTypeList.Add PT001_MSG, PT001_MSG
    publishTypeList.Add PT002_MSG, PT002_MSG
    publishTypeList.Add PT003_MSG, PT003_MSG
    publishTypeList.Add PT004_MSG, PT004_MSG
    publishTypeList.Add PT005_MSG, PT005_MSG
    
    '取次会社名
    agencyCompanyNameList.Add ACN000_MSG, ACN000_MSG
    agencyCompanyNameList.Add ACN001_MSG, ACN001_MSG
    agencyCompanyNameList.Add ACN002_MSG, ACN002_MSG
    agencyCompanyNameList.Add ACN003_MSG, ACN003_MSG
    agencyCompanyNameList.Add ACN004_MSG, ACN004_MSG
    agencyCompanyNameList.Add ACN005_MSG, ACN005_MSG
    agencyCompanyNameList.Add ACN006_MSG, ACN006_MSG
    agencyCompanyNameList.Add ACN007_MSG, ACN007_MSG
    agencyCompanyNameList.Add ACN008_MSG, ACN008_MSG
    
    '入稿グループ
    summitGroupNameList.Add SG001_MSG, SG001_MSG
    
    '取り扱いファイル形式
    serviceFileTypeList.Add SFT001_MSG, SFT001_MSG
    serviceFileTypeList.Add SFT002_MSG, SFT002_MSG
    serviceFileTypeList.Add SFT003_MSG, SFT003_MSG
    serviceFileTypeList.Add SFT004_MSG, SFT004_MSG
    serviceFileTypeList.Add SFT005_MSG, SFT005_MSG
    serviceFileTypeList.Add SFT006_MSG, SFT006_MSG
    serviceFileTypeList.Add SFT007_MSG, SFT007_MSG
    serviceFileTypeList.Add SFT008_MSG, SFT008_MSG
    serviceFileTypeList.Add SFT009_MSG, SFT009_MSG
    serviceFileTypeList.Add SFT010_MSG, SFT010_MSG
    serviceFileTypeList.Add SFT011_MSG, SFT011_MSG
    serviceFileTypeList.Add SFT012_MSG, SFT012_MSG
    serviceFileTypeList.Add SFT013_MSG, SFT013_MSG
    serviceFileTypeList.Add SFT014_MSG, SFT014_MSG
    serviceFileTypeList.Add SFT015_MSG, SFT015_MSG
    serviceFileTypeList.Add SFT016_MSG, SFT016_MSG
    serviceFileTypeList.Add SFT017_MSG, SFT017_MSG
    serviceFileTypeList.Add SFT018_MSG, SFT018_MSG
    serviceFileTypeList.Add SFT019_MSG, SFT019_MSG
    serviceFileTypeList.Add SFT020_MSG, SFT020_MSG
    serviceFileTypeList.Add SFT021_MSG, SFT021_MSG
    serviceFileTypeList.Add SFT022_MSG, SFT022_MSG
    serviceFileTypeList.Add SFT023_MSG, SFT023_MSG
    serviceFileTypeList.Add SFT024_MSG, SFT024_MSG
    serviceFileTypeList.Add SFT025_MSG, SFT025_MSG
    serviceFileTypeList.Add SFT026_MSG, SFT026_MSG
    serviceFileTypeList.Add SFT027_MSG, SFT027_MSG
    serviceFileTypeList.Add SFT028_MSG, SFT028_MSG
    serviceFileTypeList.Add SFT029_MSG, SFT029_MSG
    serviceFileTypeList.Add SFT030_MSG, SFT030_MSG
    serviceFileTypeList.Add SFT031_MSG, SFT031_MSG
    serviceFileTypeList.Add SFT032_MSG, SFT032_MSG
    serviceFileTypeList.Add SFT033_MSG, SFT033_MSG
    serviceFileTypeList.Add SFT034_MSG, SFT034_MSG
    serviceFileTypeList.Add SFT035_MSG, SFT035_MSG
    
    '承認ステータス
    proveStatusList.Add PS001_MSG, PS001_MSG
    proveStatusList.Add PS002_MSG, PS002_MSG
    proveStatusList.Add PS003_MSG, PS003_MSG
    proveStatusList.Add PS004_MSG, PS004_MSG
 
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

'取り扱いファイル形式
Public Property Get getServiceFileTypeList() As Object

       Set getServiceFileTypeList = serviceFileTypeList
        
End Property

'承認ステータス
Public Property Get getProveStatusList() As Object

       Set getProveStatusList = proveStatusList
        
End Property

Private Sub Class_Terminate()
    Set businessNameList = Nothing
    Set publishTypeList = Nothing
    Set agencyCompanyNameList = Nothing
    Set summitGroupNameList = Nothing
    Set serviceFileTypeList = Nothing
    Set proveStatusList = Nothing
End Sub

