VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Regex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iParser

Private m_RegExp As VBScript_RegExp_55.RegExp

Sub Init(RegExp As VBScript_RegExp_55.RegExp)
  Set m_RegExp = RegExp
End Sub

Function iParser_Parse(ByVal Target As String, ByVal Position As Long) As ParserOutput
  Dim RegExpMatches As MatchCollection
  Set RegExpMatches = m_RegExp.Execute(Mid(Target, Position, Len(Target)))
  
  If RegExpMatches.Count > 0 Then
    Set iParser_Parse = F.ParserOutput(True, RegExpMatches(0).Value, Position + RegExpMatches(0).Length)
  Else
    Set iParser_Parse = F.ParserOutput(False, "", Position)
  End If
End Function


