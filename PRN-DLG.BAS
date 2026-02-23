Option Explicit
Type DEVMODE    ' 68 Bytes
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
End Type

Global Const HWND_BROADCAST = -1
Global Const WM_WININICHANGE = &H1A

'Declare Function PostMessageByString% Lib "User" Alias "PostMessage" (ByVal hWnd%, ByValwMsg%, ByVal wParam%, ByVal lParam$)
Declare Function PostMessage% Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
Declare Function SendMessage& Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)


Declare Sub agCopyDataBynum Lib "Apiguide.dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount%)
Declare Function agDeviceCapabilities& Lib "Apiguide.dll" (ByVal hlib%, ByVal lpszDevice$, ByVal lpszPort$, ByVal fwCapability%, ByVal lpszOutput&, ByVal lpdm&)
Declare Function agGetAddressForObject& Lib "Apiguide.dll" (object As Any)
Declare Function agGetAddressForInteger& Lib "Apiguide.dll" Alias "agGetAddressForObject" (intnum%)
Declare Function agGetAddressForLong& Lib "Apiguide.dll" Alias "agGetAddressForObject" (intnum&)
Declare Function agGetAddressForLPSTR& Lib "Apiguide.dll" Alias "agGetAddressForObject" (ByVal lpString$)
Declare Function agGetAddressForVBString& Lib "Apiguide.dll" (vbstring$)
Declare Function agGetStringFromLPSTR$ Lib "Apiguide.dll" (ByVal lpString$)
Declare Function agExtDeviceMode% Lib "Apiguide.dll" (ByVal hwnd%, ByVal hDriver%, ByVal lpdmOutput&, ByVal lpszDevice$, ByVal lpszPort$, ByVal lpdmInput&, ByVal lpszProfile&, ByVal fwMode%)


Declare Function LoadLibrary% Lib "Kernel" (ByVal lpLibFileName$)
Declare Sub FreeLibrary Lib "Kernel" (ByVal hLibModule%)

Declare Function WritePrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal lpString$, ByVal lplFileName$)
'Declare Function WritePrivateProfileStringBynum% Lib "Kernel" Alias "WritePrivateProfileString" (ByVal lpApplicationName$, ByVal lpKeyName&, ByVal lpString&, ByVal lplFileName$)
Declare Function WriteProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal lpString$)
Declare Function GetPrivateProfileInt% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFilename$)
Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
Declare Function GetProfileInt% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%)
Declare Function GetProfileString% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%)
Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function DeleteDC% Lib "GDI" (ByVal hDC%)

Dim opis() As String, drv() As String, port() As String, papirSif%(), BinSif%()



' field selection bits
Const DM_ORIENTATION = &H1&
Const DM_PAPERSIZE = &H2&
Const DM_PAPERLENGTH = &H4&
Const DM_PAPERWIDTH = &H8&
Const DM_SCALE = &H10&
Const DM_COPIES = &H100&
Const DM_DEFAULTSOURCE = &H200&
Const DM_PRINTQUALITY = &H400&
Const DM_COLOR = &H800&
Const DM_DUPLEX = &H1000&
Const DM_YRESOLUTION = &H2000&
Const DM_TTOPTION = &H4000&

'' printer orientation selections
Const DMORIENT_PORTRAIT = 1
Const DMORIENT_LANDSCAPE = 2

'' paper selections
' Warning: The PostScript driver mistakingly uses DMPAPER_ values between
' 50 and 56.  Don't use this range when defining new paper sizes.

Const DMPAPER_LETTER = 1         'Letter 8 1/2 x 11 in
Const DMPAPER_LETTERSMALL = 2    'Letter Small 8 1/2 x 11 in
Const DMPAPER_TABLOID = 3        'Tabloid 11 x 17 in
Const DMPAPER_LEDGER = 4         'Ledger 17 x 11 in
Const DMPAPER_LEGAL = 5          'Legal 8 1/2 x 14 in
Const DMPAPER_STATEMENT = 6      'Statement 5 1/2 x 8 1/2 in
Const DMPAPER_EXECUTIVE = 7              'Executive 7 1/4 x 10 1/2 in
Const DMPAPER_A3 = 8             'A3 297 x 420 mm
Const DMPAPER_A4 = 9             'A4 210 x 297 mm
Const DMPAPER_A4SMALL = 10       'A4 Small 210 x 297 mm
Const DMPAPER_A5 = 11            'A5 148 x 210 mm
Const DMPAPER_B4 = 12            'B4 250 x 354
Const DMPAPER_B5 = 13            'B5 182 x 257 mm
Const DMPAPER_FOLIO = 14         'Folio 8 1/2 x 13 in
Const DMPAPER_QUARTO = 15        'Quarto 215 x 275 mm
Const DMPAPER_10X14 = 16         '10x14 in
Const DMPAPER_11X17 = 17         '11x17 in
Const DMPAPER_NOTE = 18          'Note 8 1/2 x 11 in
Const DMPAPER_ENV_9 = 19         'Envelope #9 3 7/8 x 8 7/8
Const DMPAPER_ENV_10 = 20        'Envelope #10 4 1/8 x 9 1/2
Const DMPAPER_ENV_11 = 21        'Envelope #11 4 1/2 x 10 3/8
Const DMPAPER_ENV_12 = 22        'Envelope #12 4 \276 x 11
Const DMPAPER_ENV_14 = 23        'Envelope #14 5 x 11 1/2
Const DMPAPER_CSHEET = 24        'C size sheet
Const DMPAPER_DSHEET = 25        'D size sheet
Const DMPAPER_ESHEET = 26        'E size sheet
Const DMPAPER_ENV_DL = 27        'Envelope DL 110 x 220mm
Const DMPAPER_ENV_C5 = 28        'Envelope C5 162 x 229 mm
Const DMPAPER_ENV_C3 = 29        'Envelope C3  324 x 458 mm
Const DMPAPER_ENV_C4 = 30        'Envelope C4  229 x 324 mm
Const DMPAPER_ENV_C6 = 31        'Envelope C6  114 x 162 mm
Const DMPAPER_ENV_C65 = 32       'Envelope C65 114 x 229 mm
Const DMPAPER_ENV_B4 = 33        'Envelope B4  250 x 353 mm
Const DMPAPER_ENV_B5 = 34        'Envelope B5  176 x 250 mm
Const DMPAPER_ENV_B6 = 35        'Envelope B6  176 x 125 mm
Const DMPAPER_ENV_ITALY = 36     'Envelope 110 x 230 mm
Const DMPAPER_ENV_MONARCH = 37   'Envelope Monarch 3.875 x 7.5 in
Const DMPAPER_ENV_PERSONAL = 38  '6 3/4 Envelope 3 5/8 x 6 1/2 in
Const DMPAPER_FANFOLD_US = 39    'US Std Fanfold 14 7/8 x 11 in
Const DMPAPER_FANFOLD_STD_GERMAN = 40  'German Std Fanfold 8 1/2 x 12 in
Const DMPAPER_FANFOLD_LGL_GERMAN = 41  'German Legal Fanfold 8 1/2 x 13 in

Const DMPAPER_USER = 256

'' printer bin selections
Const DMBIN_UPPER = 1
Const DMBIN_ONLYONE = 1
Const DMBIN_LOWER = 2
Const DMBIN_MIDDLE = 3
Const DMBIN_MANUAL = 4
Const DMBIN_ENVELOPE = 5
Const DMBIN_ENVMANUAL = 6
Const DMBIN_AUTO = 7
Const DMBIN_TRACTOR = 8
Const DMBIN_SMALLFMT = 9
Const DMBIN_LARGEFMT = 10
Const DMBIN_LARGECAPACITY = 11
Const DMBIN_CASSETTE = 14

Const DMBIN_USER = 256  'device specific bins start here

'' print qualities
Const DMRES_DRAFT = -1
Const DMRES_LOW = -2
Const DMRES_MEDIUM = -3
Const DMRES_HIGH = -4

'' color enable/disable for color printers
Const DMCOLOR_MONOCHROME = 1
Const DMCOLOR_COLOR = 2

'' Printer duplex enable
Const DMDUP_SIMPLEX = 1
Const DMDUP_VERTICAL = 2
Const DMDUP_HORIZONTAL = 3

'' TrueType options
Const DMTT_BITMAP = 1    'print TT fonts as graphics
Const DMTT_DOWNLOAD = 2  'download TT fonts as soft fonts
Const DMTT_SUBDEV = 3    'substitute device fonts for TT fonts



'' device mode function modes
Const DM_UPDATE = 1
Const DM_COPY = 2
Const DM_PROMPT = 4
Const DM_MODIFY = 8

Const DM_IN_BUFFER = 8
Const DM_IN_PROMPT = 4
Const DM_OUT_BUFFER = 2
Const DM_OUT_DEFAULT = 1

'' device capabilities indices
Const DC_FIELDS = 1
Const DC_PAPERS = 2
Const DC_PAPERSIZE = 3
Const DC_MINEXTENT = 4
Const DC_MAXEXTENT = 5
Const DC_BINS = 6
Const DC_DUPLEX = 7
Const DC_SIZE = 8
Const DC_EXTRA = 9
Const DC_VERSION = 10
Const DC_DRIVER = 11
Const DC_BINNAMES = 12
Const DC_ENUMRESOLUTIONS = 13
Const DC_FILEDEPENDENCIES = 14
Const DC_TRUETYPE = 15
Const DC_PAPERNAMES = 16
Const DC_ORIENTATION = 17
Const DC_COPIES = 18

' DC_TRUETYPE bit fields
Const DCTT_BITMAP = &H1&
Const DCTT_DOWNLOAD = &H2&
Const DCTT_SUBDEV = &H4&

Function pprinter (orient%, copies%) As Integer
On Error GoTo Printer_err
Dim default$, hwnd, opis$, port$, drv$

default = WinIniDaj("windows", "device", "")
If default <> "" Then default = Left(default, InStr(default, ",") - 1)
opis = Left(default, InStr(default, ","))
port = Mid(default, Len(opis) + 1, InStr(default, ","))
drv = Right(default, Len(opis) + Len(port) + 2)
'---------------------
    Dim dm As DEVMODE, dmout As DEVMODE
    Dim libhnd%, prhdc%, bufsize%
    Dim dminstring$, dmoutstring$, dminaddr&, dmoutaddr&
    Dim di%, ID%

    ' Load the device driver library - exit if unavailable
    libhnd% = LoadLibrary(drv + ".drv")'API LoadLibrary
    If libhnd% = 0 Then GoTo Printer_err

    ' Get a copy of the DEVMODE structure for this printer
    ' First find out how big the DEVMODE structure is
    bufsize% = agExtDeviceMode%(hwnd, libhnd%, 0, opis, port, agGetAddressForObject(dm), 0, 0)
    If bufsize <= 0 Then GoTo Printer_err
    ' Allocate a buffer of that size and get a pointer to it
    dminstring$ = String$(bufsize%, 0)
    dminaddr& = agGetAddressForVBString&(dminstring$)
    dmoutstring$ = String$(bufsize%, 0)
    dmoutaddr& = agGetAddressForVBString&(dmoutstring$)
    
    ' Get the output DEVMODE structure
    di% = agExtDeviceMode(hwnd, libhnd%, dmoutaddr&, opis, port, dminaddr&, 0, DM_OUT_BUFFER)
    
    ' Copy the data buffer into the DEVMODE structure
    agCopyDataBynum dmoutaddr&, agGetAddressForObject&(dm), 68
    ' Set the orientation, and set the dmField flag so that the function will know that it is valid.
'ShowDevMode dm   'prije
    ShowDevMode dm
    'izmjene
    dm.dmOrientation = IIf(orient, DMORIENT_PORTRAIT, DMORIENT_LANDSCAPE)
    dm.dmCopies = copies
    ''If cboPapir.ListIndex >= 0 Then dm.dmpapersize = papirSif(cboPapir.ListIndex)
    ''If cboBIN.ListIndex >= 0 Then dm.dmDefaultSource = BinSif(cboBIN.ListIndex)
    ''If cboRez <> "" Then
    ''  dm.dmPrintQuality = Val(Left(cboRez, InStr(cboRez, "x") - 2))
    ''  dm.dmYResolution = Val(Mid(cboRez, InStr(cboRez, "x") + 2))
    ''End If
    
    dm.dmFields = dm.dmFields Or DM_ORIENTATION Or DM_PAPERSIZE Or DM_DEFAULTSOURCE Or DM_PRINTQUALITY Or DM_YRESOLUTION
    ShowDevMode dm
    'upis izmjena
    agCopyDataBynum agGetAddressForObject&(dm), dmoutaddr&, bufsize
    di% = agExtDeviceMode(hwnd, libhnd%, dmoutaddr&, opis, port, dmoutaddr&, 0, 9)
    
    If libhnd% <> 0 Then FreeLibrary libhnd%
    libhnd% = 0
   'ShowDevMode dm
    pprinter = True
'---------------------
Exit Function
Printer_err:
    pprinter = False
    'greska
    Exit Function
End Function

Sub ShowDevMode (dm As DEVMODE)
    Dim crlf$
    Dim a$

    crlf$ = Chr$(13) + Chr$(10)
    a$ = "Device name = " + agGetStringFromLPSTR$(dm.dmDeviceName) + crlf$
    a$ = a$ + "Devmode Version: " + Hex$(dm.dmSpecVersion) + ", Driver version: " + Hex$(dm.dmDriverVersion) + crlf$
    a$ = a$ + "Orientation: "
    If dm.dmOrientation = DMORIENT_PORTRAIT Then a$ = a$ + "Portrait" Else a$ = a$ + "Landscape"
    a$ = a$ + crlf$
    a$ = a$ + "Field mask = " + Hex$(dm.dmFields) + crlf$
    a$ = a$ + "Copies = " + Str$(dm.dmCopies) + crlf$
    a$ = a$ + "PRINT QUALITY = " + Str$(dm.dmPrintQuality) + crlf$
    If dm.dmFields And DM_YRESOLUTION <> 0 Then
        a$ = a$ + "X,Y resolution = " + Str$(dm.dmPrintQuality) + "," + Str$(dm.dmYResolution) + crlf$
    End If
    MsgBox a$, 0, "Devmode structure"
End Sub

