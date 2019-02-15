VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapDeleteParenthesis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iMapCallback

Private Function iMapCallback_Callback(Result As Variant) As Variant
  If IsObject(Result(2)) Then
    Set iMapCallback_Callback = Result(2)
  Else
    iMapCallback_Callback = Result(2)
  End If
End Function



