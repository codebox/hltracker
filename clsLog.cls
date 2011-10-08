VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjHitLog As clsLogFile
Private mobjServerLog As clsLogFile
Private mobjDisplayLog As clsLogFile

Private mbHaveWarned As Boolean

Public Property Get ServerLogPath() As String
    ServerLogPath = App.Path & "\" & SHORT_APP_NAME & "ServerLog.log"
End Property
Public Property Get HitLogPath() As String
    HitLogPath = App.Path & "\" & SHORT_APP_NAME & "HitLog.log"
End Property
Public Property Get WindowLogPath() As String
    WindowLogPath = App.Path & "\" & SHORT_APP_NAME & "DisplayLog.log"
End Property

Public Sub NewServerInfo(objServer As clsServer)
    Dim sLogInfo As String

    If mobjSettings.DoServerLogging Then
        sLogInfo = Format$(Now, "dd/mm/yyyy hh:mm:ss") & "," & _
                   objServer.IPString & "," & _
                   CStr(objServer.Port) & "," & _
                   Replace(objServer.ServerName, ",", " ") & "," & _
                   CStr(objServer.UserCount)
                   
        WriteTextToLog sLogInfo, mobjServerLog
    End If
    
    Exit Sub
errhandler:
    ShowLogError Err.Description, ServerLogPath
    
End Sub

Public Sub NewHitInfo(sIP As String)
    Dim sLogInfo As String
    
    On Error GoTo errhandler
    
    If mobjSettings.DoHitLogging Then
        sLogInfo = Format$(Now, "dd/mm/yyyy hh:mm:ss") & "," & sIP
        WriteTextToLog sLogInfo, mobjHitLog
    End If
    
    Exit Sub
errhandler:
    ShowLogError Err.Description, HitLogPath
    
End Sub

Public Sub NewDisplayLogEntry(sEntry As String)
    On Error GoTo errhandler

    If mobjSettings.DoDisplayLogging Then
        WriteTextToLog sEntry, mobjDisplayLog
    End If
    
    Exit Sub
errhandler:
    ShowLogError Err.Description, WindowLogPath
    
End Sub

Private Sub WriteTextToLog(sText As String, objLogFile As clsLogFile)
    objLogFile.WriteData sText
End Sub

Private Sub ShowLogError(sErrDesc As String, sLogPath As String)
    Dim sErrorMsg As String
    
    If Not mbHaveWarned Then
        sErrorMsg = "There was an error when trying to update the log file '" & sLogPath & "'" & vbCrLf
        sErrorMsg = sErrorMsg & "The description of the error was:" & vbCrLf
        sErrorMsg = sErrorMsg & "    """ & sErrDesc & """" & vbCrLf & vbCrLf
        sErrorMsg = sErrorMsg & "No further messages concerning this type of error will be displayed until " & _
        sErrorMsg = sErrorMsg & APP_NAME & " is restarted, however it is unlikely the log will be updated again until" & _
        sErrorMsg = sErrorMsg & " the problem is resolved"
        MsgBox sErrorMsg, vbCritical, APP_NAME
        mbHaveWarned = True
    End If
End Sub

Private Sub Class_Initialize()
    Set mobjHitLog = New clsLogFile
    mobjHitLog.LogPath = HitLogPath
    
    
    Set mobjServerLog = New clsLogFile
    mobjServerLog.LogPath = ServerLogPath
    
    Set mobjDisplayLog = New clsLogFile
    mobjDisplayLog.LogPath = WindowLogPath
    
End Sub

Private Sub Class_Terminate()
    Set mobjHitLog = Nothing
    Set mobjServerLog = Nothing
    Set mobjDisplayLog = Nothing
End Sub
