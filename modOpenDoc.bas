Attribute VB_Name = "modOpenDoc"
Option Explicit

Public Const MAX_TARGET_LEN = 1024
Public Const ERR_OFFSET = vbObjectError + 1300

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long

Private Const SW_SHOWNORMAL = 1
Private Const SEE_MASK_INVOKEIDLIST = &HC

Private Const MOD_NAME = "modMain"

Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SW_SHOW = 5

Private Type SHELLEXECUTEINFO
        cbSize As Long
        fMask As Long
        hwnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type


Public Sub ShowProps(hParentWindow As Long, sFilePath As String)
    Dim lRetVal As Long
    Dim sTarget As String * MAX_TARGET_LEN
    Dim typINFO As SHELLEXECUTEINFO
    
    
    sTarget = sFilePath
    If sFilePath = "" Then
        Err.Raise ERR_OFFSET, , "No document was specified"
    Else
        typINFO.cbSize = LenB(typINFO)
        typINFO.fMask = SEE_MASK_INVOKEIDLIST
        typINFO.hwnd = hParentWindow
        typINFO.lpVerb = "properties"
        typINFO.lpFile = sTarget
        typINFO.nShow = SW_SHOW
        
        lRetVal = ShellExecuteEx(typINFO)
    End If
End Sub

Public Sub LaunchDoc(hParentWindow As Long, sFilePath As String)
    Dim lRetVal As Long
    Dim sTarget As String * MAX_TARGET_LEN
    Dim sErrDescription As String
    Dim nDotPosn As Integer
    Dim sDocExtension As String
    
    On Error GoTo ErrHandler
    
    If sFilePath = "" Then
        Err.Raise ERR_OFFSET, , "No document was specified"
    Else
        sTarget = sFilePath
        lRetVal = ShellExecute(hParentWindow, "Open", sTarget, "", "", SW_SHOWNORMAL)
        If lRetVal < 32 Then
          ' There was an error...
            Select Case lRetVal
                Case ERROR_FILE_NOT_FOUND
                    sErrDescription = "The file '" & sFilePath & "' could not be found."
                Case ERROR_PATH_NOT_FOUND
                    sErrDescription = "The path '" & sFilePath & "' could not be found."
                Case SE_ERR_ACCESSDENIED
                    sErrDescription = "Access to the file '" & sFilePath & "' was denied."
                Case SE_ERR_NOASSOC
                    nDotPosn = InStrRev(sFilePath, ".")
                    If nDotPosn > 0 Then
                        sDocExtension = Mid$(sFilePath, nDotPosn + 1)
                        sErrDescription = "No application is currently associated with '" & sDocExtension & "' documents."
                    Else
                        sErrDescription = "No application is currently associated with documents of this type."
                    End If
                Case SE_ERR_OOM
                    sErrDescription = "The system ran out of memory while trying to launch the document '" & sFilePath & "'"
                Case Else
                    sErrDescription = "An unexpected error occurred while trying to launch the document '" & sFilePath & "'." & vbCrLf
                    sErrDescription = sErrDescription & "The ShellExecute() return code was " & CStr(lRetVal)
            End Select
            Err.Raise ERR_OFFSET, , sErrDescription
        End If
    End If
    
    Exit Sub
ErrHandler:
   Err.Raise Err.Number, MOD_NAME & ".LaunchDoc()", Err.Description
   
End Sub

