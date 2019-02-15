VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Choice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iParser

Dim m_Parsers() As iParser

Sub Init(iParsers() As iParser)
 m_Parsers = iParsers
End Sub


Function iParser_Parse(ByVal Target As String, ByVal Position As Long) As ParserOutput
  Dim No As Long
  For No = LBound(m_Parsers) To UBound(m_Parsers)
    
    Dim Parsed As ParserOutput
    Set Parsed = m_Parsers(No).Parse(Target, Position)
    
    If Parsed.Success Then
      Set iParser_Parse = Parsed
      Exit Function
    End If
  Next
  
  Set iParser_Parse = F.ParserOutput(False, "", Position)

End Function



