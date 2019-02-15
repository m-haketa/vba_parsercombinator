VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Char"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iParser

Dim m_Dict As Scripting.Dictionary

Private Sub Class_Initialize()
  Set m_Dict = CreateObject("Scripting.Dictionary")
End Sub

Sub Init(Chars As String)
  Dim Pos As Long
  For Pos = 1 To Len(Chars)
    Dim Char As String
    Char = Mid(Chars, Pos, 1)
    
    If Not m_Dict.Exists(Char) Then
      Call m_Dict.Add(Char, Char)
    End If
  Next
End Sub


Function iParser_Parse(ByVal Target As String, ByVal Position As Long) As ParserOutput
  Dim Char As String
  Char = Mid(Target, Position, 1)
  
  If m_Dict.Exists(Char) Then
    Set iParser_Parse = F.ParserOutput(True, Char, Position + 1)
  Else
    Set iParser_Parse = F.ParserOutput(False, "", Position)
  End If
  
End Function
