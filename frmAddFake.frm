VERSION 5.00
Begin VB.Form frmAddFake 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fake Servers"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3960
      TabIndex        =   1
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   3060
      TabIndex        =   0
      Top             =   5520
      Width           =   795
   End
End
Attribute VB_Name = "frmAddFake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdOK_Click()
    On Error GoTo errhandler
    
    mobjSettings.FakeServers = mobjServerGrid.GetRegString
    mobjServerGrid.MergeServers mobjServers
    Unload Me
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdOK_Click", Err.Description, Erl
End Sub

Private Sub cmdUpdate_Click()
    Dim objServer As clsServer
    On Error GoTo errhandler
    Set objServer = BuildServer
    If Not objServer Is Nothing Then
        mobjServerGrid.UpdateServer objServer
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "cmdUpdate_Click", Err.Description, Erl
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler
    
    SetWindowPos Me.hwnd, IIf(mobjSettings.Float, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Set mobjServerGrid = New clsServerGrid
    Set mobjServerGrid.ListViewControl = lvwServers
    LoadFakeServers
    mobjServerGrid.RequestSelectionEvent
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "Form_Load", Err.Description, Erl
End Sub

Private Sub LoadFakeServers()
    Dim objServer As clsServer
    For Each objServer In mobjServers
        If objServer.ServerType = Fake Then mobjServerGrid.AddNewServer objServer
    Next objServer
End Sub

Private Sub RefreshButtons()
    Dim objServer As clsServer
    On Error GoTo errhandler
    Set objServer = mobjServerGrid.SelectedServer
    
    If Not objServer Is Nothing Then
        cmdDelete.Enabled = True
        cmdUpdate.Enabled = (txtName.Text <> objServer.ServerName) Or _
                            (txtDescription.Text <> objServer.Description) Or _
                            (txtPort.Text <> CStr(objServer.Port)) Or _
                            (txtIP1.Text & "." & txtIP2.Text & "." & txtIP3.Text & "." & txtIP4.Text <> objServer.IPString) Or _
                            (txtUsers.Text <> CStr(objServer.UserCount))
    Else
        cmdDelete.Enabled = False
        cmdUpdate.Enabled = False
    End If
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "RefreshButtons", Err.Description, Erl

End Sub

Private Function BuildServer() As clsServer
    Dim objServer As clsServer
    Dim arrIP() As String
    Dim sErrLocation As String
    Dim ctlFields As Control
    
    On Error GoTo errhandler
    
    Set objServer = BuildNewServerObject
    
    sErrLocation = "Server Name"
    Set ctlFields = txtName
    objServer.ServerName = txtName.Text
    
    sErrLocation = "Server Description"
    Set ctlFields = txtDescription
    objServer.Description = txtDescription.Text
    
    sErrLocation = "Port Number"
    Set ctlFields = txtPort
    objServer.Port = CLng(txtPort.Text)
    
    sErrLocation = "User Count"
    Set ctlFields = txtUsers
    objServer.UserCount = CLng(txtUsers.Text)
    
    sErrLocation = "IP Address"
    Set ctlFields = txtIP1
    objServer.SetIP txtIP1.Text, txtIP2.Text, txtIP3.Text, txtIP4.Text
    
    objServer.ServerType = Fake
    objServer.LastCheckedIn = Now
    Set BuildServer = objServer
    Set objServer = Nothing
    
    Exit Function
errhandler:
    If Len(sErrLocation) > 0 Then
        MsgBox "Invalid " & sErrLocation
        ctlFields.SetFocus
    Else
        ErrorReport Err.Number, MOD_NAME, "BuildServer", Err.Description, Erl
    End If
    Set BuildServer = Nothing
    
End Function

Private Sub NewServer()
    Dim objServer As clsServer
    Dim arrIP() As String
    
    On Error GoTo errhandler
    Set objServer = BuildNewServerObject
    
    objServer.SetIP "0", "0", "0", "0"
    objServer.ServerName = "New Server"
    objServer.ServerType = Fake = True
    objServer.LastCheckedIn = Now
    
    mobjServerGrid.AddNewServer objServer
    
    Set objServer = Nothing
    
    Exit Sub
errhandler:
    If Err.Number = ERR_BASE Then
        MsgBox Err.Description
    Else
        ErrorReport Err.Number, MOD_NAME, "NewServer", Err.Description, Erl
    End If
End Sub

Private Sub mobjServerGrid_SelectionChanged(objServer As clsServer)
    Dim arrIP() As String
    On Error GoTo errhandler
    
    If Not objServer Is Nothing Then
        EnableFields True
        txtName.Text = objServer.ServerName
        txtDescription.Text = objServer.Description
        txtPort = objServer.Port
        txtUsers = objServer.UserCount
        arrIP = Split(objServer.IPString, ".")
        txtIP1.Text = arrIP(0)
        txtIP2.Text = arrIP(1)
        txtIP3.Text = arrIP(2)
        txtIP4.Text = arrIP(3)
    Else
        EnableFields False
        txtName.Text = ""
        txtDescription.Text = ""
        txtPort = ""
        txtUsers = ""
        txtIP1.Text = ""
        txtIP2.Text = ""
        txtIP3.Text = ""
        txtIP4.Text = ""
    End If
    RefreshButtons
    
    Exit Sub
errhandler:
    ErrorReport Err.Number, MOD_NAME, "mobjServerGrid_SelectionChanged", Err.Description, Erl

End Sub

Private Sub EnableFields(bEnable As Boolean)
    txtName.Enabled = bEnable
    txtDescription.Enabled = bEnable
    txtIP1.Enabled = bEnable
    txtIP2.Enabled = bEnable
    txtIP3.Enabled = bEnable
    txtIP4.Enabled = bEnable
    txtPort.Enabled = bEnable
    txtUsers.Enabled = bEnable
    
    txtName.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtDescription.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtIP1.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtIP2.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtIP3.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtIP4.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtPort.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    txtUsers.BackColor = IIf(bEnable, vbWhite, DISABLED_COL)
    
End Sub

Private Sub txtDescription_Change()
    RefreshButtons
End Sub

Private Sub txtDescription_GotFocus()
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
End Sub

Private Sub txtIP1_Change()
    RefreshButtons
End Sub
Private Sub txtIP2_Change()
    RefreshButtons
End Sub
Private Sub txtIP3_Change()
    RefreshButtons
End Sub
Private Sub txtIP4_Change()
    RefreshButtons
End Sub

Private Sub txtIP1_GotFocus()
    txtIP1.SelStart = 0
    txtIP1.SelLength = Len(txtIP1.Text)
End Sub
Private Sub txtIP2_GotFocus()
    txtIP2.SelStart = 0
    txtIP2.SelLength = Len(txtIP2.Text)
End Sub
Private Sub txtIP3_GotFocus()
    txtIP3.SelStart = 0
    txtIP3.SelLength = Len(txtIP3.Text)
End Sub
Private Sub txtIP4_GotFocus()
    txtIP4.SelStart = 0
    txtIP4.SelLength = Len(txtIP4.Text)
End Sub

Private Sub txtName_Change()
    RefreshButtons
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtPort_Change()
    RefreshButtons
End Sub

Private Sub txtPort_GotFocus()

    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
End Sub

Private Sub txtUsers_Change()
    RefreshButtons
End Sub

Private Sub txtUsers_GotFocus()
    txtUsers.SelStart = 0
    txtUsers.SelLength = Len(txtUsers.Text)
End Sub
