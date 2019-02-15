VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseHoge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iParser

Function iParser_Parse(ByVal Target As String, ByVal Position As Long) As ParserOutput
  If Mid(Target, Position, 4) = "hoge" Then
    Set iParser_Parse = F.ParserOutput(True, "hoge", Position + 4)
  Else
    Set iParser_Parse = F.ParserOutput(False, "", Position)
  End If

End Function
