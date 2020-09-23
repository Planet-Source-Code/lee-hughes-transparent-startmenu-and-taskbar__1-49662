Attribute VB_Name = "mTransparentTaskbar"
'---------------------------------------------------------------------------------------
' Module     : mTransparentTaskbar
' DateTime   : 04/11/2003 ddmmyy 19:46
' Author     : Lee Hughes lphughes@btopenworld.com
' Purpose    : Makes startmenu and task bar transparent
' Unicode YN : N
' Required   : Win2000 / XP +
'---------------------------------------------------------------------------------------

'This program can be run either with a commandline
'argument or from the IDE by setting the
'Command Line Arguments property

'See:

'TransparentTaskbar Properties -
'Make -
'Command Line Arguments (transparent value 0-255)

'0 = invisible
'255 = solid

Option Explicit

Public Enum SWPEnum
  SWP_DRAWFRAME = 32
  SWP_FRAMECHANGED = 32
  SWP_HIDEWINDOW = 128
  SWP_NOACTIVATE = 16
  SWP_NOCOPYBITS = 256
  SWP_NOMOVE = 2
  SWP_NOSIZE = 1
  SWP_NOREDRAW = 8
  SWP_NOZORDER = 4
  SWP_SHOWWINDOW = 64
  SWP_NOOWNERZORDER = 512
  SWP_NOREPOSITION = 512
  SWP_NOSENDCHANGING = 1024
  SWP_DEFERERASE = 8192
  SWP_ASYNCWINDOWPOS = 16384
End Enum

Public Enum SWPhWndEnum
  HWND_BROADCAST = 65535
  HWND_BOTTOM = 1
  HWND_NOTOPMOST = -2
  HWND_TOP = 0
  HWND_TOPMOST = -1
  HWND_DESKTOP = 0
End Enum

Public Type WINDOWPOS
  hwnd As Long
  hWndInsertAfter As Long
  x As Long
  y As Long
  cX As Long
  cy As Long
  Flags As Long
End Type

Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As SWPhWndEnum, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cy As Long, ByVal wFlags As SWPEnum)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Public Enum WinGetWindowLongOffsets 'used by style and subclassing

  GWL_EXSTYLE = -20
  GWL_STYLE = -16
  GWL_WNDPROC = -4
  GWL_HINSTANCE = -6
  GWL_HWNDPARENT = -8
  GWL_ID = -12
  GWL_USERDATA = -21
End Enum
  
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal Class As String, ByVal Caption As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2

'Makes the supplied window hWnd Transparant
'Level is a value from 0 (invisible) to 255 (solid)

Private Sub MakeTrans(hwnd As Long, Level As Integer)

SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED

SetLayeredWindowAttributes hwnd, 0, Level, LWA_ALPHA

RefreshWindow hwnd

End Sub

Public Sub Main()

Dim Level As Integer
Dim Wnd As Long

If Len(Command$) Then
  Level = CInt(Command)
  
  Else
  Level = 180
  
End If

If Level < 0 Then Level = 0
If Level > 255 Then Level = 255

'Find the taskbar
Wnd = FindWindow("Shell_TrayWnd", vbNullString)
MakeTrans Wnd, Level

'Find the startmenu
Wnd = FindWindow("DV2ControlHost", vbNullString)
MakeTrans Wnd, Level


'Find IE toolbands
Wnd = FindWindow("BaseBar", vbNullString)

Do Until Wnd = 0
  
  MakeTrans Wnd, Level
  
  Wnd = GetWindow(Wnd, GW_CHILD)
  
Loop

End Sub

'After changing a windows attributes refresh the windows
'position
Public Sub RefreshWindow(hwnd As Long)
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
  
End Sub
