VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cTypes As New Collection
Public cDeclares As New Collection
Public cConsts As New Collection

Public cTypesID As Long
Public cDeclaresID As Long
Public cConstsID As Long

Private Sub Class_Initialize()
    cTypesID = 0
    cDeclaresID = 0
    cConstsID = 0
End Sub
