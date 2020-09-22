Attribute VB_Name = "basTASKBAR"
Option Explicit

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Constants for topmost.
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Enum ONTOPSETTING
    WINDOW_ONTOP = HWND_TOPMOST
    WINDOW_NOT_ONTOP = HWND_NOTOPMOST
End Enum
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Functionality to Set a window always on top or turn it off.
' Date: March,10 1999 @ 10:18:37
'------------------------------------------------------------
Public Sub SetFormOnTop(formHWND As Long, Optional sSETTING As ONTOPSETTING = WINDOW_ONTOP)
    On Error Resume Next
    Call SetWindowPos(formHWND, sSETTING, 0, 0, 0, 0, FLAGS)
End Sub

Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function
