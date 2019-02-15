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
Public Result As Variant
Public NewPosition As Long


Function toString() As String
Attribute toString.VB_UserMemId = 0
    '=Default Property
  toString = "Success:" & Success & " ,Result:" & ResultToString(Result) & " ,newPosition:" & NewPosition
End Function


'コレクションか、通常の値を受け取る前提
Function ResultToString(Parts As Variant) As String
  Dim Ret As String
  Ret = ""
   
'コレクションでない場合
  On Error GoTo NoCollection
    Dim MaxCount As Long
    MaxCount = Parts.Count
  On Error GoTo 0
    
'コレクションが空
  If MaxCount < 1 Then
    Ret = "[]"
    GoTo Finally
  End If
    
  Dim PartArray() As String
  ReDim PartArray(1 To MaxCount)
    
    
  Ret = "[ "
      
  Dim Count As Long
  Count = 1
    
  Dim Part As Variant
  For Each Part In Parts
    PartArray(Count) = ResultToString(Part)
    Count = Count + 1
  Next

  Ret = Ret & Join(PartArray, ", ") & " ]"
  
  GoTo Finally
  
NoCollection:
  If VarType(Parts) = vbString Then
    Ret = """" & Parts & """"
  Else
    Ret = Parts
  End If


Finally:
  ResultToString = Ret

End Function
