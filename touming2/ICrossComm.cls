VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICrossComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

' Implementation Interface

Public Property Get CrossTalkClass() As Object
    ' this will be used to return a cCrossComm class but since the cCrossComm class
    ' is not a public-creatable object when control is compiled to OCX we will get
    ' a compile error.  Making it Object vs cCrossComm resolves that error.
    ' However, this class when exsiting in a stand-alone usercontrol project
    ' must have its Instancing property set to 2=PublicNotCreatable
End Property


' Public Sub ReleaseMouseCapture()
    '==============
'End Sub

