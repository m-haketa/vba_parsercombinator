VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapFlattenSeqManySeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iMapCallback

Private Function iMapCallback_Callback(Result As Variant) As Variant
  Dim RetCollection As Collection
  Set RetCollection = New Collection
  
  RetCollection.Add Result(1)
  
  Dim Item As Variant
  For Each Item In Result(2)
    RetCollection.Add Item(1)
    RetCollection.Add Item(2)
  Next

  Set iMapCallback_Callback = RetCollection

End Function


