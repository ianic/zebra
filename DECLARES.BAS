Option Explicit

Declare Function SpoolFile Lib "GDI" (ByVal lpszPrinter As String, ByVal lpszPort As String, ByVal lpszJob As String, ByVal lpszFile As String) As Integer

Declare Function OSGetWindowsDirectory% Lib "Kernel" Alias "GetWindowsDirectory" (ByVal a$, ByVal b%)
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

Declare Function WriteIni% Lib "kernel" Alias "WritePrivateProfileString" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer

Declare Function SetCapture Lib "User" (ByVal hWnd As Integer) As Integer
Declare Sub ReleaseCapture Lib "User" ()

Dim WinVersion As Integer, SoundAvailable As Integer
Global VisibleFrame As Frame

Global Const TWIPS = 1
Global Const PIXELS = 3
Global Const RES_INFO = 2
Global Const MINIMIZED = 1

Type POINTAPI
   X As Integer
   y As Integer
End Type

Type Rect
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Declare Function GetWindowsDirectory Lib "Kernel" (ByVal P$, ByVal S%) As Integer
Declare Function GetSystemDirectory Lib "kernel" (ByVal P$, ByVal S%) As Integer
Declare Function GetWinFlags Lib "kernel" () As Long
Global Const WF_CPU286 = &H2&
Global Const WF_CPU386 = &H4&
Global Const WF_CPU486 = &H8&
Global Const WF_STANDARD = &H10&
Global Const WF_ENHANCED = &H20&
Global Const WF_80x87 = &H400&

Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetSystemMetrics Lib "User" (ByVal n As Integer) As Integer
Global Const SM_MOUSEPRESENT = 19

Declare Function GetDeviceCaps Lib "GDI" (ByVal hDC%, ByVal nIndex%) As Integer

Declare Function GlobalCompact Lib "kernel" (ByVal flag&) As Long
Declare Function GetFreeSpace Lib "kernel" (ByVal flag%) As Long
Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer
Global Const GFSR_SYSTEMRESOURCES = &H0
Global Const GFSR_GDIRESOURCES = &H1
Global Const GFSR_USERRESOURCES = &H2

Declare Function sndPlaySound Lib "MMSystem" (lpsound As Any, ByVal flag As Integer) As Integer
Declare Function waveOutGetNumDevs Lib "MMSystem" () As Integer

Declare Sub GetCursorPos Lib "User" (lpPoint As POINTAPI)

Declare Function TrackPopupMenuByNum% Lib "User" Alias "TrackPopupMenu" (ByVal hMenu%, ByVal wFlags%, ByVal X%, ByVal y%, ByVal nReserved%, ByVal hWnd%, ByVal lpRect&)
Declare Function GetMenu Lib "user" (ByVal hWd%) As Integer
Declare Function GetSubMenu Lib "user" (ByVal hMenu%, ByVal nPos%) As Integer
Declare Function InsertMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Global Const MF_POPUP = &H10
Global Const MF_BYPOSITION = &H400
Global Const MF_SEPARATOR = &H800

Declare Function GetDeskTopWindow Lib "User" () As Integer
Declare Function GetDC Lib "User" (ByVal hWnd%) As Integer
Declare Sub ReleaseDC Lib "User" (ByVal hWnd%, ByVal hDC%)
Declare Function BitBlt Lib "GDI" (ByVal destDC%, ByVal X%, ByVal y%, ByVal w%, ByVal h%, ByVal srchDC%, ByVal srcX%, ByVal srcY%, ByVal rop&) As Integer
Global Const SRCCOPY = &HCC0020
Global Const SRCERASE = &H440328
Global Const SRCINVERT = &H660046
Global Const SRCAND = &H8800C6

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
Global FiltPunoIme%

'rad s ini datotekom
Global Inifile$
Declare Function OSGetPrivateProfileString% Lib "Kernel" Alias "GetPrivateProfileString" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$)
Declare Function OSWritePrivateProfileString% Lib "Kernel" Alias "WritePrivateProfileString" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
   
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

'hvatanje mi�a za ispis poruke kada je iznad gumbi�a
Global capturestatus%
Global capturehandle%
Declare Function SetCapture Lib "User" (ByVal hWnd As Integer) As Integer
Declare Sub ReleaseCapture Lib "User" ()




Function DeviceColors (hDC As Integer) As Long
Const PLANES = 14
Const BITSPIXEL = 12
    DeviceColors = GetDeviceCaps(hDC, PLANES) * 2 ^ GetDeviceCaps(hDC, BITSPIXEL)
End Function

Function DosVersion ()
Dim Ver As Long, DosVer As Long
    Ver = GetVersion()
    DosVer = Ver \ &H10000
    DosVersion = Format((DosVer \ 256) + ((DosVer Mod 256) / 100), "Fixed")
End Function

Function GetSysIni (section, key)
Dim retVal As String, AppName As String, worked As Integer
    retVal = String$(255, 0)
    worked = GetPrivateProfileString(section, key, "", retVal, Len(retVal), "System.ini")
    If worked = 0 Then
	GetSysIni = "unknown"
    Else
	GetSysIni = Left(retVal, worked)
    End If
End Function

Function GetWinIni (section, key)
Dim retVal As String, AppName As String, worked As Integer
    retVal = String$(255, 0)
    worked = GetProfileString(section, key, "", retVal, Len(retVal))
    If worked = 0 Then
	GetWinIni = "unknown"
    Else
	GetWinIni = Left(retVal, worked)
    End If
End Function

Function SystemDirectory () As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    SystemDirectory = Left(WinPath, GetSystemDirectory(WinPath, Len(WinPath)))

End Function

Function WindowsDirectory () As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))
End Function

Function WindowsVersion ()
Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    WindowsVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function

