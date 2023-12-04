VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private result_ As String
Private row_ As Integer
Private col_ As Integer

Private Sub Class_Initialize()
    result_ = ""
    row_ = -1
    col_ = -1
End Sub

Property Get result() As String
    result = result_
End Property

Property Let setResult(result As String)
    result_ = result
End Property

Property Get row() As Integer
    row = row_
End Property

Property Let setRow(row As Integer)
    row_ = row
End Property

Property Get col() As Integer
    col = col_
End Property

Property Let setCol(col As Integer)
    col_ = col
End Property

Private Sub Class_Terminate()
    result_ = ""
    row_ = -1
    col_ = -1
End Sub
