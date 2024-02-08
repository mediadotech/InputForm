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
Private Const ET000_MSG As String = "必須項目が入力されてないです"
'カタカナ入力
Private Const ET001_MSG As String = "全角カタカナ入力で入力してください"
'電話番号
Private Const ET002_MSG As String = "電話番号が正しくないです"
'住所区分
Private Const ET003_MSG As String = "住所区分が正しくないです" & vbCrLf & "半角数字、1（国内）、2（国外）"
'郵便番号
Private Const ET004_MSG As String = "郵便番号が正しくないです" & vbCrLf & "半角数字、7文字"
'メールアドレス
Private Const ET005_MSG As String = "メールアドレスが正しくないです"
'発行形態
Private Const ET006_MSG As String = "発行形態が正しくないです"
'屋号
Private Const ET007_MSG As String = "屋号が正しくなりません"
'書店会社ID
Private Const ET008_MSG As String = "書店会社IDドが正しくないです" & vbCrLf & "25文字以内、半角大文字英数字、頭文字は半角大文字英字"
'サービス開始日
Private Const ET009_MSG As String = "サービス開始日が正しくなりません" & vbCrLf & "yyyy/mm/dd"
'サービス終了日
Private Const ET010_MSG As String = "サービス終了日が正しくなりません" & vbCrLf & "yyyy/mm/dd"
Private Const ET011_MSG As String = "サービス開始日より前の日付は入力できません"
'書店ID
Private Const ET012_MSG As String = "書店IDが正しくないです" & vbCrLf & "4文字以上8文字以下、半角英数字記号(記号はアンダースコア[_]のみ)"
'書店グループID
Private Const ET013_MSG As String = "書店グループIDが正しくないです" & vbCrLf & "25文字以内、半角英数字 "
'タイトルID
Private Const ET014_MSG As String = "タイトルIDが正しくないです" & vbCrLf & "半角数字"

'<支払先登録>
'支払先コード
Private Const E0100_MSG As String = "支払先登録の項目また取引先コードが入力されてないです"
Private Const E0101_MSG As String = "支払先コードは5文字にしてください"
Private Const E0102_MSG As String = "支払先コードが正しくないです" & vbCrLf & "半角英数字"
'登録番号
Private Const E0103_MSG As String = "登録番号は14文字にしてください"
Private Const E0104_MSG As String = "登録番号がただしくないです" & vbCrLf & "半角英字T＋半角数字13桁"
'追加入力フラグ
Private Const E0105_MSG As String = "発行形態追加入力フラグが正しくないです" & vbCrLf & "半角数字、0（追加入力フなし）、1（追加入力あり）"

'<出版社登録>
'支払先名
Private Const E0200_MSG As String = "出版社登録の項目また支払先名が入力されてないです"
'出版社経理コード必須フラグ
Private Const E0201_MSG As String = "出版社経理コード必須が正しくないです" & vbCrLf & "半角数字、0（必須ではない）、1（必須）"
'音声ファイル再生許諾フラグ
Private Const E0202_MSG As String = "音声ファイル再生許諾必須が正しくないです" & vbCrLf & "半角数字、0（再生許諾しない）、1（再生許諾する）"
'取次会社名
Private Const E0203_MSG As String = "取次会社名が正しくなりません"
Private Const E0204_MSG As String = "取り扱い発行形態に入力した形態だけ入力出来ます"
Private Const E0205_MSG As String = "取り扱い発行形態、約販売可能発行形態し入力した形態だけ入力出来ます"
'入稿グループ
Private Const E0206_MSG As String = "入稿グループが正しくなりません"

'<書店会社情報登録>
'書店会社ID
Private Const E0300_MSG As String = "書店会社情報の項目また書店会社IDが入力されてないです"
'取引先コード
Private Const E0301_MSG As String = "取引先コードが正しくないです" & vbCrLf & "5文字、半角英数字"
'セルシスビューア閲覧権限
Private Const E0302_MSG As String = "セルシスビューア閲覧権限が正しくないです" & vbCrLf & "半角数字、1（なし）、2（あり）"
'コンテンツDL権限
Private Const E0303_MSG As String = "コンテンツDL権限が正しくないです" & vbCrLf & "半角数字、1（なし）、2（あり）"
'コンテンツDL上限数_週
Private Const E0304_MSG As String = "コンテンツDL上限数_週が正しくないです" & vbCrLf & "1～300、半角数字"

'<書店グループ情報登録>
'書店会社ID
Private Const E0400_MSG As String = "書店グループ情報の項目また書店会社IDが入力されてないです"

'<書店情報登録>
'書店グループ名
Private Const E0500_MSG As String = "書店情報の項目また書店グループ名が入力されてないです"
'取り扱いファイル形式
Private Const E0501_MSG As String = "取り扱いファイル形式が正しくないです" & vbCrLf & "詳しい内容は5－2「入力コード一覧」を参照してください。"
'取り扱いファイル形式
Private Const E0502_MSG As String = "成人指定コンテンツ取り扱いが正しくないです" & vbCrLf & "半角数字、1（否）、2（可）"
'売上登録形式
Private Const E0503_MSG As String = "売上登録形式が正しくないです" & vbCrLf & "詳しい内容は5－2「入力コード一覧」を参照してください。"
'書店独自価格許可
Private Const E0504_MSG As String = "書店独自価格許可が正しくないです" & vbCrLf & "半角数字、1（否）、2（可）"
'予約先行公開
Private Const E0505_MSG As String = "予約先行公開が正しくないです" & vbCrLf & "半角数字、1（否）、2（可）"
'書店例外料率登録
Private Const E0506_MSG As String = "書店例外料率登録が正しくないです" & vbCrLf & "半角数字、1（必須）、2（任意）"
'帳票生成ルール
Private Const E0507_MSG As String = "帳票生成ルールが正しくないです" & vbCrLf & "半角数字、1（通常）、2（国際）"
'接続アカウント
Private Const E0508_MSG As String = "接続アカウントが正しくないです" & vbCrLf & "6文字以上25文字以下、半角英数字記号(記号はアンダースコア[_]のみ)"
'接続パスワード
Private Const E0509_MSG As String = "接続パスワードが正しくないです" & vbCrLf & "6文字以上25文字以下"

'<公開承認_書店>
'書店グループ名
Private Const E0700_MSG As String = "書店IDが入力されてないです"

'<接続先書店設定>
'書店グループ名
Private Const E0800_MSG As String = "接続先書店設定の項目また出版社名が入力されてないです"

'<タイトル種別の設定>
Private Const E1000_MSG As String = "タイトル種別の設定項目また書店グループIDが入力されてないです"
'タイトル種別ID
Private Const E1001_MSG As String = "タイトル種別IDが正しくないです" & vbCrLf & "半角英字、TS＋3文字以内"

'<公開日の設定>
Private Const E1100_MSG As String = "公開日の設定項目また書店IDが入力されてないです"
'公開自動設定フラグ
Private Const E1101_MSG As String = "公開自動設定フラグが正しくないです" & vbCrLf & "半角数字、0（しない）、1（する）"
'公開設定レーティング
Private Const E1102_MSG As String = "公開設定レーティングが正しくないです" & vbCrLf & "半角数字、1~10、20"

'<検証機能>
Private Const E1300_MSG As String = "検証機能項目またタイトルIDが入力されてないです"
'書誌情報検証
Private Const E1301_MSG As String = "書誌情報検証が正しくないです" & vbCrLf & "半角数字、0（NG）、1（OK）"
'サムネイル検証
Private Const E1302_MSG As String = "サムネイル検証が正しくないです" & vbCrLf & "半角数字、0（NG）、1（OK）"
'コンテンツファイル検証
Private Const E1303_MSG As String = "コンテンツファイル検証が正しくないです" & vbCrLf & "半角数字、0（NG）、1（OK）"

'<レーティング>
Private Const E1400_MSG As String = "レーティング設定項目またタイトルIDが入力されてないです"
'レーティング
Private Const E1401_MSG As String = "レーティングが正しくないです" & vbCrLf & "半角数字、1~10、20、99(確認中)"

'<公開日自動設定>
Private Const E1500_MSG As String = "公開日自動設定項目またタイトルIDが入力されてないです"
'コンテンツID
Private Const E1501_MSG As String = "コンテンツIDが正しくないです" & vbCrLf & "半角英数字、BT0+17文字"
'承認ステータス
Private Const E1502_MSG As String = "承認ステータスが正しくないです" & vbCrLf & "半角数字、0（未設定）、1（承認）、9（保留）、99（削除）"


'共通
Public Property Get ET000() As String
    ET000 = ET000_MSG
End Property

Public Property Get ET001() As String
    ET001 = ET001_MSG
End Property

Public Property Get ET002() As String
    ET002 = ET002_MSG
End Property

Public Property Get ET003() As String
    ET003 = ET003_MSG
End Property

Public Property Get ET004() As String
    ET004 = ET004_MSG
End Property

Public Property Get ET005() As String
    ET005 = ET005_MSG
End Property

Public Property Get ET006() As String
    ET006 = ET006_MSG
End Property

Public Property Get ET007() As String
    ET007 = ET007_MSG
End Property

Public Property Get ET008() As String
    ET008 = ET008_MSG
End Property

Public Property Get ET009() As String
    ET009 = ET009_MSG
End Property

Public Property Get ET010() As String
    ET010 = ET010_MSG
End Property

Public Property Get ET011() As String
    ET011 = ET011_MSG
End Property

Public Property Get ET012() As String
    ET012 = ET012_MSG
End Property

Public Property Get ET013() As String
    ET013 = ET013_MSG
End Property

Public Property Get ET014() As String
    ET014 = ET014_MSG
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

'書店グループ情報登録
Public Property Get E0400() As String
    E0400 = E0400_MSG
End Property
    
'書店情報登録
Public Property Get E0500() As String
    E0500 = E0500_MSG
End Property

Public Property Get E0501() As String
    E0501 = E0501_MSG
End Property

Public Property Get E0502() As String
    E0502 = E0502_MSG
End Property

Public Property Get E0503() As String
    E0503 = E0503_MSG
End Property

Public Property Get E0504() As String
    E0504 = E0504_MSG
End Property

Public Property Get E0505() As String
    E0505 = E0505_MSG
End Property

Public Property Get E0506() As String
    E0506 = E0506_MSG
End Property

Public Property Get E0507() As String
    E0507 = E0507_MSG
End Property

Public Property Get E0508() As String
    E0508 = E0508_MSG
End Property

Public Property Get E0509() As String
    E0509 = E0509_MSG
End Property


'<公開承認_書店>
Public Property Get E0700() As String
    E0700 = E0700_MSG
End Property

'<接続先書店設定>
Public Property Get E0800() As String
    E0800 = E0800_MSG
End Property

'<タイトル種別の設定>
Public Property Get E1000() As String
    E1000 = E1000_MSG
End Property

Public Property Get E1001() As String
    E1001 = E1001_MSG
End Property

'<公開日の設定>
Public Property Get E1100() As String
    E1100 = E1100_MSG
End Property

Public Property Get E1101() As String
    E1101 = E1101_MSG
End Property

Public Property Get E1102() As String
    E1102 = E1102_MSG
End Property

'<検証機能>
Public Property Get E1300() As String
    E1300 = E1300_MSG
End Property

Public Property Get E1301() As String
    E1301 = E1301_MSG
End Property

Public Property Get E1302() As String
    E1302 = E1302_MSG
End Property

Public Property Get E1303() As String
    E1303 = E1303_MSG
End Property

'<レーティング>
Public Property Get E1400() As String
    E1400 = E1400_MSG
End Property

Public Property Get E1401() As String
    E1401 = E1401_MSG
End Property

'<公開日自動設定>
Public Property Get E1500() As String
    E1500 = E1500_MSG
End Property

Public Property Get E1501() As String
    E1501 = E1501_MSG
End Property

Public Property Get E1502() As String
    E1502 = E1502_MSG
End Property
