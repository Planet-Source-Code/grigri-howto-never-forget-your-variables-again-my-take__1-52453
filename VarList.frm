VERSION 5.00
Begin VB.Form VarList 
   BorderStyle     =   0  'None
   Caption         =   "Variable List"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstVars 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "VarList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_HIDEWINDOW As Long = &H80

Private Const HWND_TOPMOST As Long = -1

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCaretPos Lib "user32.dll" (lpPoint As Any) As Long
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, lpPoint As Any) As Long

Private sSelectedEntry As String
Private bExitOnBlur As Boolean
Private bExitLoop As Boolean

Public Function ShowVariables(cm As CodeModule, ByVal hWndParent As Long) As String
    Dim mbrs As Members, mbr As Member
    Dim tw As Single
    
    bExitOnBlur = False
    
    Set mbrs = cm.Members
    
    ' Load the form, but don't display it
    Me.Visible = False
    
    ' Populate the listbox with the members
    lstVars.Clear
    For Each mbr In mbrs
        If mbr.Type = vbext_mt_Variable Or mbr.Type = vbext_mt_Const Then
            If mbr.Scope = vbext_Private Then
                ' Add the item
                lstVars.AddItem mbr.Name
                ' Get the maximum text width
                ' note that this is in pixels (ScaleMode property)
                If TextWidth(mbr.Name) > tw Then tw = TextWidth(mbr.Name)
            End If
        End If
    Next
    
    ' Retrieve the coordinates of the caret
    ' This is in client (pixel) coordinates relative to the VbaWindow
    Dim pos(1 To 2) As Long
    GetCaretPos pos(1)
    ' Offset left a bit
    pos(1) = pos(1) + 5
    ' Translate this to screen coordinates
    ClientToScreen hWndParent, pos(1)
    
    'SetParent hwnd, hWndParent
    'SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_CHILD
    ' Move the window and make it topmost, size it, and display it
    SetWindowPos hwnd, HWND_TOPMOST, pos(1), pos(2), 0, 0, SWP_NOSIZE
    Me.Width = (tw + 20) * Screen.TwipsPerPixelX
    Me.Visible = True
    
    ' Let the message system keep up
    DoEvents
    
    ' Give us the focus
    SetFocusAPI hwnd
    ' Give the listbox the focus
    lstVars.SetFocus
    
    ' Setup member variables
    bExitLoop = False
    sSelectedEntry = ""
    bExitOnBlur = True
    
    ' Loop until we need to stop
    Do
        DoEvents
    Loop Until bExitLoop = True
    
    bExitOnBlur = False
    
    ' Set the focus back to the code window
    SetFocusAPI hWndParent
    
    ' Hide the window immediately
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_HIDEWINDOW
    'Me.Visible = False
    ' Set the old style and parent back
    'SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) And Not (WS_CHILD)
    'SetParent hwnd, 0
    ' Unload the form
    Unload Me
    
    ' Return the string that was selected
    ShowVariables = sSelectedEntry
End Function

Private Sub Form_GotFocus()
    lstVars.SetFocus
End Sub

Private Sub Form_Load()
    ' Copy the list's font object
    ' this is necessary to measure the
    ' text strings easily
    Font = lstVars.Font
End Sub

Private Sub Form_Resize()
    lstVars.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub lstVars_DblClick()
    If lstVars.ListIndex = -1 Then Exit Sub
    ' User double-clicked an item - save selection and exit
    sSelectedEntry = lstVars.List(lstVars.ListIndex)
    bExitLoop = True
End Sub

Private Sub lstVars_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then
        ' Enter or space pressed - save selection and exit
        KeyAscii = 0
        sSelectedEntry = lstVars.List(lstVars.ListIndex)
        bExitLoop = True
    ElseIf KeyAscii = 27 Then
        ' Escape key pressed - exit
        KeyAscii = 0
        bExitLoop = True
    End If
End Sub

Private Sub lstVars_LostFocus()
    If bExitOnBlur Then
        ' Exit
        bExitLoop = True
    End If
End Sub


