VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iParser

Private m_Parser As iParser
Private m_MapCallback As iMapCallback

Sub Init(iParser As iParser, iMapCallback As iMapCallback)
  Set m_Parser = iParser
  Set m_MapCallback = iMapCallback
End Sub

'Sub SetCallback(Callback As iMapCallback)
'  Set m_Callback = Callback
'End Sub

Function iParser_Parse(ByVal Target As String, ByVal Position As Long) As ParserOutput
  Dim Ret As ParserOutput
  
  Set Ret = m_Parser.Parse(Target, Position)
  
  If Ret.Success Then
  
    Set iParser_Parse = F.ParserOutput(Ret.Success, m_MapCallback.Callback(Ret.Result), Ret.NewPosition)
    
  Else
    Set iParser_Parse = Ret
  End If
  
  
End Function


