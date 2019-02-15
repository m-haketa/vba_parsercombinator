VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapParsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iMapCallback

Private Function iMapCallback_Callback(Result As Variant) As Variant
  iMapCallback_Callback = Result + "‚Æ‚¢‚¤•¶š—ñ‚ğƒp[ƒX‚Å‚«‚½‚æ"
End Function
