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

'‰®†
Private Const C0100_MSG As String = "‚È‚µ"
Private Const C0101_MSG As String = "Š”®‰ïĞ`"
Private Const C0102_MSG As String = "`Š”®‰ïĞ"
Private Const C0103_MSG As String = "—LŒÀ‰ïĞ`"
Private Const C0104_MSG As String = "`—LŒÀ‰ïĞ"
Private Const C0105_MSG As String = "‡“¯‰ïĞ`"
Private Const C0106_MSG As String = "à’c–@l`"
Private Const C0107_MSG As String = "ˆê”ÊĞ’c–@l`"
Private Const C0108_MSG As String = "Œö‰và’c–@l`"
Private Const C0109_MSG As String = "Ğ’c–@l`"
Private Const C0110_MSG As String = "ŠwZ–@l`"
Private Const C0111_MSG As String = "Œö‰vĞ’c–@l`"
Private Const C0112_MSG As String = "ˆê”Êà’c–@l`"
Private Const C0113_MSG As String = "`‡‘‰ïĞ"
Private Const C0114_MSG As String = "`‡“¯‰ïĞ"
Private Const C0115_MSG As String = "‡‘‰ïĞ`"
Private Const C0116_MSG As String = "“Á’è”ñ‰c—˜Šˆ“®–@l`"
Private Const C0117_MSG As String = "‘—§‘åŠw–@l`"

Private Sub Class_Initialize()
    Set businessNameList = CreateObject("Scripting.Dictionary")
    
    businessNameList.Add C0100_MSG, C0100_MSG
    businessNameList.Add C0101_MSG, C0101_MSG
    businessNameList.Add C0102_MSG, C0102_MSG
    businessNameList.Add C0103_MSG, C0103_MSG
    businessNameList.Add C0104_MSG, C0104_MSG
    businessNameList.Add C0105_MSG, C0105_MSG
    businessNameList.Add C0106_MSG, C0106_MSG
    businessNameList.Add C0107_MSG, C0107_MSG
    businessNameList.Add C0108_MSG, C0108_MSG
    businessNameList.Add C0109_MSG, C0109_MSG
    businessNameList.Add C0110_MSG, C0110_MSG
    businessNameList.Add C0111_MSG, C0111_MSG
    businessNameList.Add C0112_MSG, C0112_MSG
    businessNameList.Add C0113_MSG, C0113_MSG
    businessNameList.Add C0114_MSG, C0114_MSG
    businessNameList.Add C0115_MSG, C0115_MSG
    businessNameList.Add C0116_MSG, C0116_MSG
    businessNameList.Add C0117_MSG, C0117_MSG
End Sub

'‰®†‚ÌƒŠƒXƒg‚ğæ“¾
Public Property Get getBusinessNameList() As Object

    Set getBusinessNameList = businessNameList
        
End Property

Private Sub Class_Terminate()
    Set businessNameList = Nothing
End Sub

