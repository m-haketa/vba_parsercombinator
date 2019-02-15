VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapReduce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iMapCallback

Private Function iMapCallback_Callback(Result As Variant) As Variant
  Set iMapCallback_Callback = AddtoCollection(Result, New Collection)
End Function


'今のところ、不完全
'RetCollectionも書き換わるので注意。
'返り値とRetCollectionは同じCollectionを指す
Private Function AddtoCollection(ByRef Items As Variant, ByRef RetCollection As Collection) As Collection

'コレクションでない場合
  On Error GoTo NoCollection
    Dim MaxCount As Long
    MaxCount = Items.Count
  On Error GoTo 0

'コレクションが空
  If MaxCount < 1 Then
    GoTo Finally
  End If

  Dim Item As Variant
  For Each Item In Items
    Call AddtoCollection(Item, RetCollection)
  Next

  GoTo Finally

NoCollection:
  Call RetCollection.Add(Items)

Finally:
  Set AddtoCollection = RetCollection

End Function
