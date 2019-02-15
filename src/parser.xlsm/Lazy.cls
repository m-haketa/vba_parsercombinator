VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lazy"
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
  Set iParser_Parse = m_Parser.Parse(Target, Position)
End Function

