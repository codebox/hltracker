VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMirroredTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Address As String
Public Enabled As Boolean

Public Function Clone() As clsMirroredTracker
    Dim objClone As clsMirroredTracker
    Set objClone = New clsMirroredTracker
    objClone.Address = Address
    objClone.Enabled = Enabled
    Set Clone = objClone
    Set objClone = Nothing
End Function
