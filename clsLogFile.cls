VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LogPath As String

Private Const REDUCE_BY = 1024& * 1024
Private Const CHUNK_COPY_SIZE = 1024

Public Sub WriteData(sText As String)
    Dim lCurrentLogSize As Long
    Dim lReadPointer As Long
    Dim lWritePointer As Long
    
    
    Open LogPath For Append As #2
    Print #2, sText
    Close #2
    
End Sub



