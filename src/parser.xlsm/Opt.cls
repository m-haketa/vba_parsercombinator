VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Opt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iParser

Dim m_Parser As iParser

Sub Init(iParser As iParser)
  Set m_Parser = iParser
End Sub


Function iParser_Parse(ByVal Target As String, ByVal Position As Long) As ParserOutput
  Dim Parsed As ParserOutput
  Set Parsed = m_Parser.Parse(Target, Position)
    
  If Parsed.Success Then
    Set iParser_Parse = Parsed
  Else
    Set iParser_Parse = F.ParserOutput(True, "", Position)
  End If
  
End Function
