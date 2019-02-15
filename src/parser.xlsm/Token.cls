VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iParser

Private m_Token As String
Private m_Length As Long

Sub Init(Str As String)
  m_Token = Str
  m_Length = Len(m_Token)
End Sub

Function iParser_Parse(ByVal Target As String, ByVal Position As Long) As ParserOutput
  If Mid(Target, Position, m_Length) = m_Token Then
    Set iParser_Parse = F.ParserOutput(True, m_Token, Position + m_Length)
  Else
    Set iParser_Parse = F.ParserOutput(False, "", Position)
  End If
End Function

