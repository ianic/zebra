VERSION 2.00
Begin Form frmPrnOdabir 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Podešavanje printera"
   ClientHeight    =   4815
   ClientLeft      =   255
   ClientTop       =   1875
   ClientWidth     =   6390
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Height          =   5325
   Left            =   195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6390
   Top             =   1425
   Width           =   6510
   Begin SSPanel Panel3D2 
      Align           =   2  'Align Bottom
      BevelInner      =   1  'Inset
      BorderWidth     =   1
      Font3D          =   1  'Raised w/light shading
      ForeColor       =   &H000000FF&
      Height          =   1110
      Left            =   0
      TabIndex        =   1
      Top             =   3705
      Width           =   6390
      Begin SSCommand cmdOdustani 
         Caption         =   "&Odustani"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   795
         Left            =   3480
         Picture         =   PRN_DLG.FRX:0000
         TabIndex        =   2
         Top             =   150
         Width           =   1500
      End
      Begin SSCommand cmdPrihvati 
         Caption         =   "&Prihvati"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   795
         Left            =   1350
         Picture         =   PRN_DLG.FRX:01A8
         TabIndex        =   0
         Top             =   150
         Width           =   1500
      End
   End
   Begin SSPanel Panel3D1 
      Align           =   1  'Align Top
      BevelInner      =   1  'Inset
      BorderWidth     =   1
      Font3D          =   1  'Raised w/light shading
      ForeColor       =   &H000000FF&
      Height          =   3675
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6390
      Begin SSFrame Frame3D1 
         Caption         =   "Odabir pisaèa"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Width           =   6255
         Begin SSPanel pa 
            BevelInner      =   1  'Inset
            BorderWidth     =   1
            Font3D          =   1  'Raised w/light shading
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   90
            TabIndex        =   17
            Top             =   360
            Width           =   6105
            Begin ComboBox cboPrn 
               BackColor       =   &H00C0C0C0&
               FontBold        =   -1  'True
               FontItalic      =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   9.75
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               ForeColor       =   &H00800000&
               Height          =   360
               Left            =   30
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   30
               Width           =   6015
            End
         End
      End
      Begin SSFrame Frame3D2 
         Caption         =   "Podešavanje pisaèa"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   2655
         Left            =   60
         TabIndex        =   5
         Top             =   930
         Width           =   6255
         Begin TextBox edtCopies 
            Height          =   285
            Left            =   1440
            TabIndex        =   19
            Text            =   "1"
            Top             =   2250
            Width           =   735
         End
         Begin ComboBox cboBin 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1080
            Width           =   3975
         End
         Begin ComboBox cboRez 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   720
            Width           =   3975
         End
         Begin ComboBox cboPapir 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   3975
         End
         Begin SSOption optOrj 
            Caption         =   "Položena stranica"
            ForeColor       =   &H00800000&
            Height          =   192
            Index           =   1
            Left            =   2190
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1890
            Width           =   2052
         End
         Begin SSOption optOrj 
            Caption         =   "Vertikalno (portret)"
            ForeColor       =   &H00800000&
            Height          =   228
            Index           =   0
            Left            =   2190
            TabIndex        =   7
            Top             =   1500
            Width           =   2052
         End
         Begin SSPanel p 
            BackColor       =   &H00FFFFFF&
            Font3D          =   1  'Raised w/light shading
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   4470
            TabIndex        =   6
            Top             =   1500
            Width           =   135
            Begin Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   870
               X2              =   990
               Y1              =   30
               Y2              =   180
            End
         End
         Begin Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Broj kopija"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   330
            TabIndex        =   18
            Top             =   2280
            Width           =   915
         End
         Begin Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Izvor papira"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   15
            Top             =   1140
            Width           =   1020
         End
         Begin Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rezolucija"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   14
            Top             =   810
            Width           =   900
         End
         Begin Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Velièina papira"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   13
            Top             =   450
            Width           =   1275
         End
         Begin Label l 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Orijenatcija stranice"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   12
            Top             =   1500
            Width           =   1710
         End
      End
   End
End
Option Explicit
Dim opis() As String, drv() As String, port() As String, papirSif%(), BinSif%()
'' Printer and device support constants
Dim printer_iCopies%
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

Sub cboPrn_Click ()
   On Error Resume Next
   : DajPapire: cboPapir.ListIndex = 0
   : DajRezolucije: cboRez.ListIndex = 0
   : DajBinNames: cboBIN.ListIndex = 0
   ReadDef
End Sub

Sub cmdOdustani_Click ()
   Unload Me
End Sub

Sub cmdPrihvati_Click ()
  If IsNumeric(edtCopies.Text) Then printer_iCopies = CInt(edtCopies.Text)
   'If mode = 1 Then
   '   Hide
   'Else
      UpisWIN_ini
      Unload Me
   'End If
   
   'Dim X%
   'x = WriteProfileString(Chr$(0), Chr$(0), Chr$(0))
   'x = PostMessageByString(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
   'X = PostMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")'"windows")
End Sub

Sub Command1_Click ()
   printer.EndDoc
   printer.Print "TEST"
   printer.NewPage
   printer.EndDoc

End Sub

Sub Command3_Click ()
   Form_Load
End Sub

Sub DajBinNames ()
    Dim ID%, kaj%, libhnd%, X%
    Dim BrojPodataka%, Podatak$
kaj = DC_BINNAMES
   
   If cboPrn.ListIndex < 0 Then Exit Sub' MsgBox "ODABERI PRINTER!": Exit Sub
   ID = cboPrn.ListIndex + 1
    
    libhnd% = LoadLibrary(drv(ID) + ".drv")' Load the device driver library - exit if unavailable
    If libhnd% = 0 Then Exit Sub
       BrojPodataka% = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, 0, 0)
       If BrojPodataka% <> 0 Then 'ako postoji
	  Podatak$ = String$(24 * BrojPodataka%, 0)   ' Now dimension the string large enough to hold them all
	  X = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, agGetAddressForVBString&(Podatak$), 0)
       End If
      
      cboBIN.Clear
      For X% = 1 To BrojPodataka%
	 cboBIN.AddItem agGetStringFromLPSTR$(Mid$(Podatak$, (X% - 1) * 24 + 1))
      Next X%

    FreeLibrary (libhnd%)

DajBinSifre '!!!!!!!!!!!!!

End Sub

Sub DajBinSifre ()
   Dim ID%, kaj%, libhnd%, X%
   Dim BrojPodataka%', polje%()
kaj = DC_BINS
   
   If cboPrn.ListIndex < 0 Then Exit Sub' MsgBox "ODABERI PRINTER!": Exit Sub
   ID = cboPrn.ListIndex + 1
    
    libhnd% = LoadLibrary(drv(ID) + ".drv")
       If libhnd% = 0 Then Exit Sub
       BrojPodataka% = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, 0, 0)
       If BrojPodataka% <> 0 Then 'ako postoji
	  ReDim BinSif(BrojPodataka)
	  X = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, agGetAddressForInteger&(BinSif(0)), 0)
       End If
    FreeLibrary (libhnd%)
End Sub

Sub DajPapire ()
    Dim ID%, kaj%, libhnd%, X%
    Dim BrojPodataka%, Podatak$
kaj = DC_PAPERNAMES
   
   If cboPrn.ListIndex < 0 Then Exit Sub' MsgBox "ODABERI PRINTER!": Exit Sub
   ID = cboPrn.ListIndex + 1
    
    libhnd% = LoadLibrary(drv(ID) + ".drv")' Load the device driver library - exit if unavailable
    If libhnd% = 0 Then Exit Sub
      BrojPodataka% = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, 0, 0)
      If BrojPodataka% <> 0 Then 'ako postoji
	 Podatak$ = String$(64 * BrojPodataka%, 0)   ' Now dimension the string large enough to hold them all
	 X = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, agGetAddressForVBString&(Podatak$), 0)
      End If
      
      cboPapir.Clear
      For X% = 1 To BrojPodataka%
	 cboPapir.AddItem agGetStringFromLPSTR$(Mid$(Podatak$, (X% - 1) * 64 + 1))
      Next X%
    
    FreeLibrary (libhnd%)

DajPapireSifre '!!!!!!!!!!!!!

End Sub

Sub DajPapireSifre ()
   Dim ID%, kaj%, libhnd%, X%
   Dim BrojPodataka%', polje%()
kaj = DC_PAPERS
   
   If cboPrn.ListIndex < 0 Then Exit Sub' MsgBox "ODABERI PRINTER!": Exit Sub
   ID = cboPrn.ListIndex + 1
    
    libhnd% = LoadLibrary(drv(ID) + ".drv")
       If libhnd% = 0 Then Exit Sub
       BrojPodataka% = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, 0, 0)
       If BrojPodataka% <> 0 Then 'ako postoji
	  ReDim papirSif(BrojPodataka)
	  X = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, agGetAddressForInteger&(papirSif(0)), 0)
       End If
    FreeLibrary (libhnd%)
End Sub

Sub DajPrnInfo ()
   Dim broj%', opis() As String, drv() As String, port() As String
   Dim linija$, sstr$, po$, to1$, to2$, X%, tmp$, def%, default$
   ReDim opis(15), drv(15), port(15)
   broj = 0

   'text1 = "": DoEvents
   tmp = Space(255)
   X = GetWindowsDirectory(tmp$, 255)
   tmp = Trim(tmp): tmp = Left(tmp, Len(tmp) - 1)
   
   Open tmp & "\win.ini" For Input As #1
      X = False
      Do While Not EOF(1)
	 If Not X Then 'traži [PrinterPorts]
	    Input #1, linija
	    If InStr(linija, "[PrinterPorts]") Then X = True
	 Else 'poèeo [PrinterPorts]
	    Input #1, linija, po, to1, to2
	    If InStr(linija, "[") Or InStr(po, "[") Or InStr(to1, "[") Or InStr(to2, "[") Then
	       X = False 'kraj
	    Else
	       broj = broj + 1
	       If InStr(linija, "=") <> 0 Then
		   opis(broj) = Left(linija, InStr(linija, "=") - 1)
		   drv(broj) = Mid(linija, InStr(linija, "=") + 1)
	       End If
	       port(broj) = po
	    End If
	 End If
      Loop
   Close #1 ' Close file.
   'device=HP LaserJet IIP,HPPCL,LPT3:
   default = WinIniDaj("windows", "device", ""): If default <> "" Then default = Left(default, InStr(default, ",") - 1)
   def = 0
   
   cboPrn.Clear
   For X = 1 To broj
      If default = opis(X) Then
	 def = X
      End If
      cboPrn.AddItem opis(X)
      cboPrn.ItemData(cboPrn.NewIndex) = X
   Next
   If def > 0 Then cboPrn.ListIndex = def - 1


End Sub

Sub DajRezolucije ()
   Dim ID%, kaj%, libhnd%, X%
   Dim BrojPodataka%, polje()  As Long
kaj = DC_ENUMRESOLUTIONS
   
   If cboPrn.ListIndex < 0 Then Exit Sub' MsgBox "ODABERI PRINTER!": Exit Sub
   ID = cboPrn.ListIndex + 1
    
   libhnd% = LoadLibrary(drv(ID) + ".drv")
   If libhnd% = 0 Then Exit Sub
      BrojPodataka% = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, 0, 0)
      
      If BrojPodataka% > -1 Then 'ako postoji
	 ReDim polje(1, BrojPodataka)
	 X = agDeviceCapabilities(libhnd%, opis$(ID), port$(ID), kaj%, agGetAddressForLong&(polje&(0, 0)), 0)
      End If
      cboRez.Clear
      For X% = 0 To BrojPodataka% - 1
	 'Debug.Print polje(0, x) & ":" & polje(1, x)
	 cboRez.AddItem polje(0, X) & " x " & polje(1, X)
      Next X%
   FreeLibrary (libhnd%)
End Sub

Sub Form_Load ()
'If mode = 1 Then
'   Height = 2700
'Else
'   Height = 6060
'End If
Dim i%
MakniItalic Me
CenterForm Me
   DajPrnInfo
   DoEvents
   ReadDef
End Sub

Sub MsgBlast1_Message (msgVal As Integer, wParam As Integer, lParam As Long, ReturnVal As Long)
    MsgBox Str(msgVal)
End Sub

Sub optOrj_Click (index As Integer, value As Integer)
   Dim i%
   'i = IIf(index = 0, 1, 0)
   'p(i).Visible = value
   'DoEvents
   'p(index).Visible = Not p(i).Visible
   'If Index = 0 Then p(1).Visible = value: p(0).Visible = Not value
   'If Index = 1 Then p(0).Visible = value: p(1).Visible = Not value
   'If index = 0 Then p(0).Width = 500 Else p(0).Width = 1500
   p(0).Width = IIf(index = 0, 500, 800)
   p(0).Height = IIf(index = 0, 700, 500)
   p(0).Left = IIf(index = 0, 5000, 4850)
   p(0).Top = IIf(index = 0, 1700, 1800)
End Sub

Function PostaviCbo (c As ComboBox, polje%(), value)
   Dim i%
   Dim t()
   PostaviCbo = False
   For i = 0 To UBound(polje, 1)' - 1
      If polje(i) = value Then
	 PostaviCbo = True
	 c.ListIndex = i: Exit For
      End If
   Next
End Function

Function PostaviCbo2 (c As ComboBox, value1, value2)
   Dim i%
   'On Error Resume Next
   PostaviCbo2 = False
   For i = 0 To c.ListCount - 1
      If c.List(i) = value1 & " x " & value2 Then
	 PostaviCbo2 = True
	 c.ListIndex = i: Exit For
      End If
   Next
End Function

Sub ReadDef ()
   Dim ID%
   If cboPrn.ListIndex < 0 Then Exit Sub' MsgBox "ODABERI PRINTER!": Exit Sub
   ID = cboPrn.ListIndex + 1
    
    Dim dm As DEVMODE, dmout As DEVMODE
    Dim libhnd%, prhdc%, bufsize%
    Dim dminstring$, dmoutstring$, dminaddr&, dmoutaddr&
    Dim di%

    ' Load the device driver library - exit if unavailable
    libhnd% = LoadLibrary(drv(ID) + ".drv") 'API LoadLibrary
    If libhnd% = 0 Then GoTo cleanup2

    ' Get a copy of the DEVMODE structure for this printer
    ' First find out how big the DEVMODE structure is
    bufsize% = agExtDeviceMode%(hWnd, libhnd%, 0, opis$(ID), port$(ID), agGetAddressForObject(dm), 0, 0)
    If bufsize <= 0 Then MsgBox "GREŠKA!", 16: Exit Sub
    ' Allocate a buffer of that size and get a pointer to it
    dminstring$ = String$(bufsize%, 0)
    dminaddr& = agGetAddressForVBString&(dminstring$)
    dmoutstring$ = String$(bufsize%, 0)
    dmoutaddr& = agGetAddressForVBString&(dmoutstring$)
    
    ' Get the output DEVMODE structure
    di% = agExtDeviceMode(hWnd, libhnd%, dmoutaddr&, opis$(ID), port$(ID), dminaddr&, 0, DM_OUT_BUFFER)
    
    ' Copy the data buffer into the DEVMODE structure
    agCopyDataBynum dmoutaddr&, agGetAddressForObject&(dm), 68
    ' Set the orientation, and set the dmField flag so that the function will know that it is valid.

'ShowDevMode dm   'prije
   On Error Resume Next
   If Not PostaviCbo(cboPapir, papirSif(), dm.dmpapersize) Then
      'MsgBox "greška kod odabira papira br. " & dm.dmpapersize
   End If
   If Not PostaviCbo(cboBIN, BinSif(), dm.dmDefaultSource) Then
      'MsgBox "greška kod odabira uvlaèenja papira br. " & dm.dmDefaultSource
   End If
   If Not PostaviCbo2(cboRez, dm.dmPrintQuality, dm.dmYResolution) Then
      'If cboRez.ListCount > 0 Then MsgBox "greška kod odabira rezolucije " & dm.dmPrintQuality & " x " & dm.dmYResolution
   End If
   If dm.dmOrientation = DMORIENT_PORTRAIT Then optOrj(0) = True Else optOrj(1) = True
   
   'uzmi broj kopija
   edtCopies.Text = Format$(dm.dmCopies)
   printer_iCopies = dm.dmCopies

    If libhnd% <> 0 Then FreeLibrary libhnd%
    libhnd% = 0

Exit Sub
    
cleanup2:
    If prhdc% <> 0 Then di% = DeleteDC%(prhdc%)
    If libhnd% <> 0 Then FreeLibrary libhnd%


End Sub

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

Sub UpisWIN_ini ()
   Dim ID%
   If cboPrn.ListIndex < 0 Then Exit Sub' MsgBox "ODABERI PRINTER!": Exit Sub
   ID = cboPrn.ListIndex + 1
    
    Dim dm As DEVMODE, dmout As DEVMODE
    Dim libhnd%, prhdc%, bufsize%
    Dim dminstring$, dmoutstring$, dminaddr&, dmoutaddr&
    Dim di%

    ' Load the device driver library - exit if unavailable
    libhnd% = LoadLibrary(drv(ID) + ".drv") 'API LoadLibrary
    If libhnd% = 0 Then GoTo cleanup3

    ' Get a copy of the DEVMODE structure for this printer
    ' First find out how big the DEVMODE structure is
    bufsize% = agExtDeviceMode%(hWnd, libhnd%, 0, opis$(ID), port$(ID), agGetAddressForObject(dm), 0, 0)
    If bufsize <= 0 Then MsgBox "GREŠKA!", 16: Exit Sub
    ' Allocate a buffer of that size and get a pointer to it
    dminstring$ = String$(bufsize%, 0)
    dminaddr& = agGetAddressForVBString&(dminstring$)
    dmoutstring$ = String$(bufsize%, 0)
    dmoutaddr& = agGetAddressForVBString&(dmoutstring$)
    
    ' Get the output DEVMODE structure
    di% = agExtDeviceMode(hWnd, libhnd%, dmoutaddr&, opis$(ID), port$(ID), dminaddr&, 0, DM_OUT_BUFFER)
    
    ' Copy the data buffer into the DEVMODE structure
    agCopyDataBynum dmoutaddr&, agGetAddressForObject&(dm), 68
    ' Set the orientation, and set the dmField flag so that the function will know that it is valid.
'ShowDevMode dm   'prije
    
    'izmjene
    dm.dmOrientation = IIf(optOrj(0), DMORIENT_PORTRAIT, DMORIENT_LANDSCAPE)
    If cboPapir.ListIndex >= 0 Then dm.dmpapersize = papirSif(cboPapir.ListIndex)
    If cboBIN.ListIndex >= 0 Then dm.dmDefaultSource = BinSif(cboBIN.ListIndex)
    If cboRez <> "" Then
      dm.dmPrintQuality = Val(Left(cboRez, InStr(cboRez, "x") - 2))
      dm.dmYResolution = Val(Mid(cboRez, InStr(cboRez, "x") + 2))
    End If
    dm.dmFields = dm.dmFields Or DM_ORIENTATION Or DM_PAPERSIZE Or DM_DEFAULTSOURCE Or DM_PRINTQUALITY Or DM_YRESOLUTION Or DM_COPIES
    
    'upisi broj kopija
    If IsNumeric(edtCopies.Text) Then
      dm.dmCopies = CInt(edtCopies.Text)
      printer_iCopies = CInt(edtCopies.Text)
    End If
    

    'upis izmjena
    agCopyDataBynum agGetAddressForObject&(dm), dmoutaddr&, bufsize
    di% = agExtDeviceMode(hWnd, libhnd%, dmoutaddr&, opis$(ID), port$(ID), dmoutaddr&, 0, 9)
    
    If libhnd% <> 0 Then FreeLibrary libhnd%
    libhnd% = 0
   'ShowDevMode dm
   
   'upis default printera
   WinIniPisi "windows", "device", opis(ID) & "," & drv(ID) & "," & port(ID)

Exit Sub
    
cleanup3:
    If prhdc% <> 0 Then di% = DeleteDC%(prhdc%)
    If libhnd% <> 0 Then FreeLibrary libhnd%


End Sub

Function WinIniDaj (s1$, s2$, s3$)
   Dim X%, temp$
   temp = Space(255)
   'X = GetProfileString(s1, s2, s3, temp, Len(temp), IniFile)
   X = GetProfileString(s1, s2, s3, temp, Len(temp))
   temp = Trim(temp)
   WinIniDaj = Left(temp, Len(temp) - 1)
End Function

Sub WinIniPisi (s1$, s2$, s3$)
   Dim X%
   X = WriteProfileString(s1$, s2$, s3$)
End Sub

