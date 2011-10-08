VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    SetWindowPos Me.hwnd, IIf(mobjSettings.Float, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
End Sub

Private Sub fraOptions_Click()
    Unload Me

End Sub

Private Sub Label1_Click()
    Unload Me

End Sub

Private Sub Label2_Click()
    Unload Me

End Sub

Private Sub lblVersion_Click()
    Unload Me

End Sub

Private Sub Picture1_Click()
    Unload Me

End Sub

Private Sub txtAbout_Change()
    Unload Me

End Sub
