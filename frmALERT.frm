VERSION 5.00
Begin VB.Form frmALERT 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUNLOAD 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   240
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1680
      Top             =   240
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New PSC project(s) have been posted"
      ForeColor       =   &H0080FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image picBOX 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1560
      Picture         =   "frmALERT.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   435
   End
End
Attribute VB_Name = "frmALERT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Reverse As Boolean
Private Sub Form_Load()
    On Error Resume Next
    With Me
        SetFormOnTop .hwnd, WINDOW_ONTOP
        .Reverse = False
        .Top = Screen.Height - GetTaskbarHeight
        .Left = Screen.Width - .ScaleWidth - 500
        .Height = 50
        Me.Refresh
        Me.Timer.Enabled = True
    End With
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    With Me
        .picBOX.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
End Sub

Private Sub lblMSG_Click()
    On Error Resume Next
    frmMAIN.GotoPSC
End Sub

Private Sub picBOX_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub Timer_Timer()
    On Error Resume Next
    If Me.Reverse = False Then
        Me.Height = Me.Height + 100
        Me.Top = Screen.Height - GetTaskbarHeight - Me.Height
        Me.picBOX.Refresh
        If Me.Height >= 750 Then
            Me.Timer.Enabled = False
            Me.tmrUNLOAD.Enabled = True
        End If
    Else
        Me.Height = Me.Height - 100
        Me.Top = Screen.Height - GetTaskbarHeight - Me.Height
        If Me.Height <= 50 Then
            Me.Timer.Enabled = False
            Me.tmrUNLOAD.Enabled = False
            Unload Me
        End If
    End If
End Sub
Private Sub tmrUnload_Timer()
    On Error Resume Next
    Me.Reverse = True
    Me.Timer.Enabled = True
End Sub

