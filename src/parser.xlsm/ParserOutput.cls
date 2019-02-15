VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Success As Boolean
Public Result As String
Public NewPosition As Long


Function toString() As String
  toString = "Success:" & Success & " ,Result:" & Result & " ,newPosition:" & NewPosition
End Function
