VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "systray.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC Alert"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraACTIONS 
      Caption         =   "Actions:"
      Height          =   975
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   3135
      Begin VB.CommandButton cmdPSC 
         Caption         =   "..."
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdNOW 
         Caption         =   "..."
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblLABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Go to Planet Source Code."
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblLABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Check for updates."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame fraOPTIONS 
      Caption         =   "Options:"
      Height          =   975
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtMIN 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkSOUND 
         Alignment       =   1  'Right Justify
         Caption         =   "Sound alert"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label lblLABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Check every"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "min(s)"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin SHDocVwCtl.WebBrowser webCTL 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      ExtentX         =   4895
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   3720
      Top             =   3360
   End
   Begin SysTrayCtl.cSysTray Tray 
      Left            =   4800
      Top             =   3360
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "frmMAIN.frx":08CA
      TrayTip         =   "PSC Alert"
   End
   Begin VB.Timer IconTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   3360
   End
   Begin VB.Menu mnuRIGHT 
      Caption         =   "RightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuCONFIGURE 
         Caption         =   "Confi&gure..."
      End
      Begin VB.Menu mnuPSC 
         Caption         =   "&Goto PSC"
      End
      Begin VB.Menu mnuCHECK 
         Caption         =   "C&heck now"
      End
      Begin VB.Menu mnuCANCEL 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnuSEP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEXIT 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' PSC Alert (VB Submitted Code)
' This is very simple
' in nature.  The PSC Alert application will monitor
' the scrolling newest uploads list for VB on the
' Planet Source Code site and notify you when a
' new upload has appeared. PSC Alert runs in the
' System Tray and the icon will flash when a new
' upload is detected. Simply double click on the
' tray icon to go to the newest uploads page. The
' ZIP dones NOT contain the SYSTRAY.OCX which you will need
' but since PSC does not allow OCX's (which I agree with)
' you can find it at ftp://24.7.173.85/systray.ocx
'
' The code is very simple and I am sure many of you can
' customize it to provide more functionality, like
' monitor other areas of PSC and to set user definded
' times to check. I just wanted to post this like
' this while it is still simple in nature. Feel
' free to email with any questions: lafeverc@saic.com
' day lafeverc@home.com night USA East Coast
' http://lafever.iscool.net
'------------------------------------------------------------
Option Explicit
Public LastSource
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Private SafeToClose As Boolean
Public MaxMin As Long
Public Sub ExecuteLink(LINK As String)
    On Error Resume Next
    Dim lRet As Long
    If LINK <> "" Then
        lRet = ShellExecute(0, "open", LINK, "", App.Path, SW_SHOWNORMAL)
        If lRet >= 0 And lRet <= 32 Then
            MsgBox "Error jumping to:" & LINK, 48, "Warning"
        End If
    End If
End Sub
Private Sub chkSOUND_Click()
    On Error Resume Next
    REG.SaveSetting "Settings", "Sound", Me.chkSOUND.Value
End Sub
Private Sub cmdNOW_Click()
    On Error Resume Next
    CheckForUpdate
End Sub
Private Sub cmdPSC_Click()
    On Error Resume Next
    GotoPSC
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Dim psc As String
    Me.Visible = False
    Set REG = New CREG
    With REG
        .Company = "LTECH"
        .AppName = "PSCAlert"
        Me.MaxMin = CLng(.GetSetting("Settings", "Minutes", 1))
        Me.chkSOUND.Value = .GetSetting("Settings", "Sound", 1)
        Me.txtMIN.Text = Me.MaxMin
    End With
    SafeToClose = False
    LastSource = ""
    psc = LoadResString(102)
    With webCTL
        webCTL.Navigate psc
    End With
    Set Me.Tray.TrayIcon = LoadResPicture(103, vbResIcon)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If Me.Visible = True Then
        If SafeToClose = False Then
            Cancel = True
            Me.Visible = False
        End If
    End If
End Sub
Private Sub Form_Resize()
    On Error Resume Next
End Sub
Private Sub IconTimer_Timer()
    On Error Resume Next
    Static iState As Long
    If iState = 0 Then
        Set Me.Tray.TrayIcon = LoadResPicture(101, vbResIcon)
        iState = 1
    Else
        Set Me.Tray.TrayIcon = LoadResPicture(102, vbResIcon)
        iState = 0
    End If
End Sub
Private Sub mnuCHECK_Click()
    On Error Resume Next
    CheckForUpdate
End Sub
Private Sub mnuCONFIGURE_Click()
    On Error Resume Next
    Me.Visible = True
End Sub
Private Sub mnuEXIT_Click()
    On Error Resume Next
    SafeToClose = True
    Unload Me
End Sub
Private Sub mnuPSC_Click()
    On Error Resume Next
    GotoPSC
End Sub
Private Sub Timer_Timer()
    On Error Resume Next
    Static CurMin As Long
    If CurMin = 0 Then
        CheckForUpdate
        CurMin = Me.MaxMin
    Else
        CurMin = CurMin - 1
    End If
End Sub
Private Sub Tray_MouseDblClick(Button As Integer, Id As Long)
    On Error Resume Next
    GotoPSC
End Sub
Private Sub Tray_MouseUp(Button As Integer, Id As Long)
    On Error Resume Next
    If Button = vbRightButton Then
        Me.PopupMenu mnuRIGHT, , , , mnuPSC
    End If
End Sub
Private Sub txtMIN_Change()
    On Error Resume Next
    If txtMIN.Text = "" Then
        txtMIN.Text = REG.GetSetting("Settings", "Minutes", 1)
    Else
        If IsNumeric(txtMIN.Text) Then
            If CLng(txtMIN.Text) <= 0 Then
                txtMIN.Text = 1
            End If
            REG.SaveSetting "Settings", "Minutes", CLng(txtMIN.Text)
            txtMIN.SelStart = Len(txtMIN.Text)
            MaxMin = REG.GetSetting("Settings", "Minutes", 1)
        Else
            txtMIN.Text = REG.GetSetting("Settings", "Minutes", 1)
            txtMIN.SelStart = Len(txtMIN.Text)
        End If
    End If
End Sub
Private Sub webCTL_DownloadComplete()
    On Error Resume Next
    Me.Timer.Enabled = True
    Set Me.Tray.TrayIcon = LoadResPicture(102, vbResIcon)
    If webCTL.Document.body.InnerHTML <> LastSource Then
        frmALERT.Visible = True
        LastSource = webCTL.Document.body.InnerHTML
        Me.Timer.Enabled = False
        Me.IconTimer.Enabled = True
        If Me.chkSOUND.Value = 1 Then
            PlayWaveRes appsnd_ALERT
        End If
        Me.Timer.Enabled = True
    End If
End Sub
Public Sub GotoPSC()
    On Error GoTo ErrorGotoPSC
    Me.IconTimer.Enabled = False
    Set Me.Tray.TrayIcon = LoadResPicture(102, vbResIcon)
    ExecuteLink LoadResString(101)
    Exit Sub
ErrorGotoPSC:
    MsgBox Err & ":Error in GotoPSC.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub CheckForUpdate()
    On Error Resume Next
    Me.Timer.Enabled = False
    Set Me.Tray.TrayIcon = LoadResPicture(103, vbResIcon)
    Me.webCTL.Refresh
End Sub
