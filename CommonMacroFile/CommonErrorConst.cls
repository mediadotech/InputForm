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

'<共通>
'必須項目
Private Const E0000_MSG As String = "必須項目を入力してください"
'カタカナ入力
Private Const E0001_MSG As String = "入力が正しくありません" & vbCrLf & "全角カタカナ"
'電話番号
Private Const E0002_MSG As String = "番号が正しくありません" & vbCrLf & "（ハイフン（-）込みの半角数字）"
'住所区分
Private Const E0003_MSG As String = "住所区分が正しくありません" & vbCrLf & "半角数字、1（国内）、2（国外）"
'郵便番号
Private Const E0004_MSG As String = "郵便番号が正しくありません" & vbCrLf & "7文字、半角数字"
'メールアドレス
Private Const E0005_MSG As String = "メールアドレスが正しくありません" & vbCrLf & "半角、@が含まれてるメール形式"
'発行形態
Private Const E0006_MSG As String = "発行形態が正しくありません"
'屋号
Private Const E0007_MSG As String = "屋号が正しくなりません"

'<支払先登録>
'支払先コード
Private Const E0100_MSG As String = "支払先登録の項目また取引先コードが入力されてません"
Private Const E0101_MSG As String = "支払先コードは5文字にしてください"
Private Const E0102_MSG As String = "支払先コードは半角英数字です"
'登録番号
Private Const E0103_MSG As String = "登録番号は14文字にしてください"
Private Const E0104_MSG As String = "登録番号は半角英数字です"
Private Const E0105_MSG As String = "登録番号形式が違います" & vbCrLf & "半角英字T＋半角数字13桁"
'追加入力フラグ
Private Const E0106_MSG As String = "発行形態追加入力フラグが正しくありません" & vbCrLf & "半角数字、0（追加入力フなし）、1（追加入力あり）"

'<出版社登録>
'支払先名
Private Const E0200_MSG As String = "出版社登録の項目また支払先名が入力されてません"
'出版社経理コード必須フラグ
Private Const E0201_MSG As String = "出版社経理コード必須が正しくありません" & vbCrLf & "半角数字、0（必須ではない）、1（必須）"
'音声ファイル再生許諾フラグ
Private Const E0202_MSG As String = "音声ファイル再生許諾必須が正しくありません" & vbCrLf & "半角数字、0（再生許諾しない）、1（再生許諾する）"
'取次会社名
Private Const E0203_MSG As String = "取次会社名が正しくなりません"
Private Const E0204_MSG As String = "取り扱い発行形態に入力した形態だけ入力出来ます"
Private Const E0205_MSG As String = "取り扱い発行形態、約販売可能発行形態し入力した形態だけ入力出来ます"
'入稿グループ
Private Const E0206_MSG As String = "入稿グループが正しくなりません"
'サービス開始日
Private Const E0207_MSG As String = "サービス開始日が正しくなりません" & vbCrLf & "yyyy/mm/dd"
'サービス終了日
Private Const E0208_MSG As String = "サービス終了日が正しくなりません" & vbCrLf & "yyyy/mm/dd"
Private Const E0209_MSG As String = "サービス開始日より前の日付は入力できません"

'<書店会社情報登録>
'書店会社ID
Private Const E0300_MSG As String = "書店会社情報の項目また書店会社IDが入力されてません"
Private Const E0301_MSG As String = "書店会社IDドが正しくありません" & vbCrLf & "25文字以内、半角大文字英数字、頭文字は半角大文字英字"
'取引先コード
Private Const E0302_MSG As String = "取引先コードが正しくありません" & vbCrLf & "5文字、半角英数字"
'セルシスビューア閲覧権限
Private Const E0303_MSG As String = "セルシスビューア閲覧権限が正しくありません" & vbCrLf & "半角数字、0（なし）、1（あり）"
'コンテンツDL権限
Private Const E0304_MSG As String = "コンテンツDL権限が正しくありません" & vbCrLf & "半角数字、0（なし）、1（あり）"
'コンテンツDL上限数_週
Private Const E0305_MSG As String = "コンテンツDL上限数_週が正しくありません" & vbCrLf & "1～300、半角数字"


'共通
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


'支払先登録
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


'出版社登録
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

'書店会社情報登録
Public Property Get E0300() As String
    E0300 = E0300_MSG
End Property

Public Property Get E0301() As String
    E0301 = E0301_MSG
End Property

Public Property Get E0302() As String
    E0302 = E0302_MSG
End Property

Public Property Get E0303() As String
    E0303 = E0303_MSG
End Property

Public Property Get E0304() As String
    E0304 = E0304_MSG
End Property

Public Property Get E0305() As String
    E0305 = E0305_MSG
End Property
