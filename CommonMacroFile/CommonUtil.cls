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

'全角確認
Public Function hasZenkaku(s As String) As Boolean
  If Len(s) <> LenB(StrConv(s, vbFromUnicode)) Then
    hasZenkaku = True
  Else
    hasZenkaku = False
  End If
End Function

'文字列に日本語が含まれているか確認
Public Function findJapanese(str As String) As String
    '正規表現クラスオブジェクト
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    '検索条件＝日本語以外を抽出
    reg.Pattern = "[^ぁ-んァ-ヶ一-龠〃々〆〇｡-ﾟ]"
    '文字列の最後まで検索する
    reg.Global = True
    '指定セルの日本語以外を空文字に置き換える
    findJapanese = reg.Replace(str, "")
      Set reg = Nothing
End Function

'文字列が半角数字になっているか確認
Public Function checkNarrowNum(str As String) As String
    '正規表現クラスオブジェクト
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    '検索条件＝半角数字
    reg.Pattern = "[0-9]"
    '文字列の最後まで検索する
    reg.Global = True
    '指定セルの半角数字を空文字に置き換える
    checkNarrowNum = reg.Replace(str, "")
    Set reg = Nothing
End Function

'文字列が半角大英文字になっているか確認
Public Function checkUpperChar(str As String) As String
    '正規表現クラスオブジェクト
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    '検索条件＝半角大英文字
    reg.Pattern = "[A-Z]"
    '文字列の最後まで検索する
    reg.Global = True
    '指定セルの半角大英文字を空文字に置き換える
    checkNarrowNum = reg.Replace(str, "")
    Set reg = Nothing
End Function


'文字列が半角大文字英数字になっているか確認
Public Function checkNarrowUpperCharNum(str As String) As String
    Dim targetStr As String
    '指定セルの半角数字を空文字に置き換える
    targetStr = checkNarrowNum(str)
    '指定セルの半角大英文字を空文字に置き換える
    targetStr = checkUpperChar(targetStr)
    
    checkNarrowUpperCharNum = targetStr
    
End Function

'文字列がカタカナになっているか確認
Public Function checkKatakana(str As String) As String
    '正規表現クラスオブジェクト
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    '検索条件＝カタカナ
    reg.Pattern = "[ァ-ヴ\―\ー\－]"
    '文字列の最後まで検索する
    reg.Global = True
    '指定セルのカタカナを空文字に置き換える
    checkKatakana = reg.Replace(str, "")
      Set reg = Nothing
End Function

'文字列が半角数字とハイフンになっているか確認
Public Function checkNarrowNumAndHyphen(str As String) As String
    '正規表現クラスオブジェクト
    Dim reg
    Set reg = CreateObject("VBScript.RegExp")
    '検索条件＝半角数字
    reg.Pattern = "[0-9\-]"
    '文字列の最後まで検索する
    reg.Global = True
    '指定セルの半角数字を空文字に置き換える
    checkNarrowNumAndHyphen = reg.Replace(str, "")
      Set reg = Nothing
End Function

'正規表現で値がメールアドレスの形式を確認する
Private Function checkMailAdess(getVal As Variant) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    ' メールアドレス形式の正規表現
    reg.Pattern = "^[a-zA-Z0-9]+([a-zA-Z0-9\._-]){0,63}@[a-zA-Z0-9-]+\.(?:[\.a-zA-Z0-9-]+)*$"
   '文字列の最後まで検索する
    reg.Global = True
   '大文字と小文字を区別(True：しない、Flase：する)
    reg.IgnoreCase = True
    checkMailAdess = reg.Test(getVal)
    Set reg = Nothing
End Function

'フラグ（0,1系）の確認
Public Function checkFlg(row As Integer, col As Integer) As String
    checkFlg = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '全角数字かつ0また1以外の数字がある
    If Len(checkNarrowNum(targetStr)) > 0 Or Cells(row, col).Value < 0 Or Cells(row, col).Value > 1 Then
        checkFlg = "NG"
        Exit Function
    End If
End Function

'屋号確認
Public Function findBusinessName(str As String) As Boolean
    findBusinessName = commonStr.getBusinessNameList.Exists(str)
End Function

'発行形態確認
Public Function findPublishType(str As String) As Boolean
    findPublishType = commonStr.getPublishTypeList.Exists(str)
End Function

'取次会社名確認
Public Function findAgencyCompanyName(str As String) As Boolean
    findAgencyCompanyName = commonStr.getAgencyCompanyNameList.Exists(str)
End Function

'入稿グループ確認
Public Function findSummitGroupName(str As String) As Boolean
    findSummitGroupName = commonStr.getSummitGroupNameList.Exists(str)
End Function
    
'全角カタカナ確認
Public Function checkWideKana(row As Integer, col As Integer) As String
    checkWideKana = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '半角かつカタカナではない
    If Not hasZenkaku(targetStr) Or Len(checkKatakana(targetStr)) > 0 Then
        checkWideKana = "E0001"
        Exit Function
    End If
End Function

'屋号確認
Public Function checkBusinessName(row As Integer, col As Integer) As String
    checkBusinessName = "OK"
    Dim businessName As String
    businessName = Cells(row, col).Value
    If Not findBusinessName(businessName) Then
        checkBusinessName = "E0007"
        Exit Function
    End If
End Function

'電話番号確認
Public Function checkNumberWithHyphen(row As Integer, col As Integer) As String
    checkNumberWithHyphen = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    'ハイフンが含まれてるのかかる数字とハイフン以外の文字がある
    If InStr(targetStr, "-") = 0 Or Len(checkNarrowNumAndHyphen(targetStr)) > 0 Then
        checkNumberWithHyphen = "E0002"
        Exit Function
    End If
End Function

'メールアドレス確認
Public Function checkMailAdressList(row As Integer, col As Integer) As String
    checkMailAdressList = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '全角文字が含まれている
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

'発行形態項目確認
Public Function checkPublishTypeList(row As Integer, col As Integer) As String
    checkPublishTypeList = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '全角文字が含まれている
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

 '住所区分からの項目共通処理
Public Function addressDivisionCheck(row As Integer, col As Integer) As Object
    
    Dim targetResult As InputResult
    Set targetResult = New InputResult
    
    Dim targetRow As Integer
    targetRow = row
    
    targetResult.setRow = targetRow
    
    '住所区分
    Dim adressDivision As Integer
    adressDivision = col
    '郵便番号
    Dim zipCode As Integer
    zipCode = adressDivision + 1
    '住所
    Dim adress As Integer
    adress = adressDivision + 2
     
     
    '住所区分
    targetResult.setCol = adressDivision
    targetResult.setResult = checkFlg(targetRow, adressDivision)
    If Not targetResult.result = "OK" Then
        targetResult.setResult = "E0003"
        Set addressDivisionCheck = targetResult
        Exit Function
    End If

    '郵便番号
    targetResult.setCol = zipCode
    targetResult.setResult = checkZipCode(targetRow, zipCode)
    If Not targetResult.result = "OK" Then
       Set addressDivisionCheck = targetResult
       Exit Function
    End If
  
    '住所
    targetResult.setCol = adress
    If IsEmpty(Cells(targetRow, adress).Value) = True Then
        targetResult.setResult = ""
        Set addressDivisionCheck = targetResult
        Exit Function
    End If
    
    '項目に問題なし
    targetResult.setResult = "OK"
    Set addressDivisionCheck = targetResult
    
End Function

'郵便番号確認
Private Function checkZipCode(row As Integer, col As Integer) As String
    checkZipCode = "OK"
    Dim targetStr As String
    targetStr = Cells(row, col).Value
    '半角数字ではないかつ7文字ではない
    If Len(checkNarrowNum(targetStr)) > 0 Or Not Len(targetStr) = 7 Then
        checkZipCode = "E0004"
        Exit Function
    End If
End Function

Private Sub Class_Terminate()
    Set commonStr = Nothing
    Set errorStr = Nothing
End Sub
