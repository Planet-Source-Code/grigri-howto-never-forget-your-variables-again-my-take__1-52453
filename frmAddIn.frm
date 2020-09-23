VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Add In"
   ClientHeight    =   1110
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Private Sub Command1_Click()
    If CaptainHook.IsHooked Then
        CaptainHook.StopHook
    Else
        CaptainHook.SetHook
    End If
    UpdateControls
End Sub

Private Sub Form_Load()
    UpdateControls
End Sub

Private Sub UpdateControls()
    If CaptainHook.IsHooked Then
        Label1.Caption = "Hook is running"
        Command1.Caption = "Stop Hook"
    Else
        Label1.Caption = "Hook is not running"
        Command1.Caption = "Start Hook"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CaptainHook.IsHooked Then CaptainHook.StopHook
End Sub
