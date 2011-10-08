VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFilterRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CheckForMatch(objServer As clsServer) As Boolean
End Function

Public Property Get UniqueID() As String
End Property

Public Property Get Description() As String
End Property

Public Sub BuildFromString(sData As String)
End Sub

Public Function WriteToString() As String
End Function

Public Function Clone() As IFilterRule
End Function

Public Property Get Name() As String
End Property

Public Property Get Block() As Boolean
End Property

Public Property Get Valid() As Boolean
End Property

Public Property Get Enabled() As Boolean
End Property

Public Property Let Enabled(bEnabled As Boolean)
End Property

