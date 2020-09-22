Attribute VB_Name = "APICalls"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'###
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
''''
Public Type POINTAPI
  x As Long
  y As Long
End Type
''''
Public Type WINDOWPLACEMENT
  Length As Long
  Flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type
'###
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
''''
Public Const VK_LBUTTON = &H1
Public Const PS_SOLID = 0
Public Const R2_NOTXORPEN = 10
Public Const BLACK_PEN = 7
''''
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
''''
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const VS_FF_DEBUG = &H1&
''''
Public Const WS_OVERLAPPED = &H0
Public Const WS_POPUP = &H80000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_EX_MDICHILD = &H40
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_CONTEXTHELP = &H400
Public Const WS_EX_RIGHT = &H1000
Public Const WS_EX_LEFT = &H0
Public Const WS_EX_RTLREADING = &H2000
Public Const WS_EX_LTRREADING = &H0
Public Const WS_EX_LEFTSCROLLBAR = &H4000
Public Const WS_EX_RIGHTSCROLLBAR = &H0
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE + WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE + WS_EX_TOOLWINDOW + WS_EX_TOPMOST)
