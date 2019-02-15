VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Seq"
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
  Dim Result As Collection
  Set Result = New Collection
  
  Dim NewPosition As Long
  NewPosition = Position
  
  Dim No As Long
  For No = LBound(m_Parsers) To UBound(m_Parsers)
    
    Dim Parsed As ParserOutput
    Set Parsed = m_Parsers(No).Parse(Target, NewPosition)
    
    If Parsed.Success Then
      Call Result.Add(Parsed.Result)
      NewPosition = Parsed.NewPosition
    Else
      Set iParser_Parse = F.ParserOutput(False, Result, Position)
      Exit Function
    End If
  Next
  
  Set iParser_Parse = F.ParserOutput(True, Result, NewPosition)

End Function




