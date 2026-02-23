VERSION 2.00
Begin Form postavke 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Postavke"
   ClientHeight    =   5505
   ClientLeft      =   4590
   ClientTop       =   990
   ClientWidth     =   4410
   ControlBox      =   0   'False
   Height          =   6015
   Left            =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4410
   Top             =   540
   Width           =   4530
   Begin SSCommand cmdKorbaz 
      Caption         =   "Korigiraj bazu"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   405
      Left            =   90
      TabIndex        =   20
      Top             =   4350
      Width           =   1635
   End
   Begin SSFrame Frame3D1 
      Caption         =   "Backup direktorij"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   675
      Left            =   90
      TabIndex        =   17
      Top             =   3630
      Width           =   4215
      Begin SSCommand cmdDIR 
         Caption         =   "DIR"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   3210
         TabIndex        =   19
         Top             =   210
         Visible         =   0   'False
         Width           =   495
      End
      Begin TextBox txtbackup 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   240
         Width           =   4065
      End
   End
   Begin SSCommand cmdINI 
      Caption         =   "INI file"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   3540
      TabIndex        =   15
      Top             =   1530
      Width           =   735
   End
   Begin SSCheck chk2 
      Caption         =   "ispis svih podataka na raèunu"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   150
      TabIndex        =   14
      Top             =   1530
      Value           =   -1  'True
      Width           =   2505
   End
   Begin SSCommand dp 
      Caption         =   "DEFAULT pisaè"
      Enabled         =   0   'False
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   585
      Left            =   120
      TabIndex        =   13
      Top             =   570
      Width           =   1575
   End
   Begin SSCheck chk1 
      Caption         =   "DRAFT ispisi"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   150
      TabIndex        =   12
      Top             =   60
      Value           =   -1  'True
      Width           =   1545
   End
   Begin SSCommand CPrihvati 
      Caption         =   "&Upiši"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   720
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1155
   End
   Begin SSCommand CZaboravi 
      Caption         =   "&Odustani"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   2640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1155
   End
   Begin SSFrame fr 
      Caption         =   "printer port kod DRAFT ispisa"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1455
      Index           =   1
      Left            =   1770
      TabIndex        =   3
      Top             =   30
      Width           =   2505
      Begin SSOption optPrinterPort 
         Caption         =   "LPT 1"
         Font3D          =   3  'Inset w/light shading
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   540
         Width           =   885
      End
      Begin SSOption optPrinterPort 
         Caption         =   "LPT 2"
         Font3D          =   3  'Inset w/light shading
         Height          =   195
         Index           =   2
         Left            =   1050
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   540
         Width           =   1155
      End
      Begin SSOption optPrinterPort 
         Caption         =   "Ispis na ekran"
         Font3D          =   3  'Inset w/light shading
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "100"
         Top             =   1050
         Width           =   1605
      End
      Begin SSOption optPrinterPort 
         Caption         =   "LPT 3"
         Font3D          =   3  'Inset w/light shading
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   780
         Width           =   885
      End
      Begin SSOption optPrinterPort 
         Caption         =   "LPT 4"
         Font3D          =   3  'Inset w/light shading
         Height          =   195
         Index           =   4
         Left            =   1050
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "4"
         Top             =   750
         Width           =   1215
      End
      Begin SSOption optPrinterPort 
         Caption         =   "PRN"
         Font3D          =   3  'Inset w/light shading
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Tag             =   "0"
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin SSFrame fr 
      Caption         =   "Boje"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1545
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   2010
      Width           =   4215
      Begin SSPanel pb 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         BorderWidth     =   0
         Caption         =   "Tablice"
         Font3D          =   3  'Inset w/light shading
         Height          =   315
         Index           =   2
         Left            =   2730
         TabIndex        =   16
         Top             =   390
         Width           =   1275
      End
      Begin SSPanel pb 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         BorderWidth     =   0
         Caption         =   "Izlazni raèuni"
         Font3D          =   3  'Inset w/light shading
         Height          =   1065
         Index           =   1
         Left            =   1410
         TabIndex        =   0
         Top             =   390
         Width           =   1275
      End
      Begin SSPanel pb 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         BorderWidth     =   0
         Caption         =   "Ulazni raèuni"
         Font3D          =   3  'Inset w/light shading
         Height          =   1065
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   390
         Width           =   1245
      End
   End
   Begin CommonDialog cmd 
      Left            =   2070
      Top             =   4860
   End
End
Option Explicit

Sub chk1_Click (Value As Integer)
fr(1).Enabled = Value
dp.Enabled = Not Value
End Sub

Sub cmdDIR_Click ()
cmd.InitDir = txtBackup
cmd.Action = 1
txtBackup = cmd.Filename
End Sub

Sub cmdINI_Click ()
Dim X
X = Shell("Notepad.exe " & inifile, 1)
End Sub

Sub cmdKorbaz_Click ()
Dim ssql$
mp 11
'On Error Resume Next
ssql = "INSERT INTO KIRA ( rb, br_racun, NAZIV, por_br, vsp, OSN, pdv, datum, kosn, kpdv, adresa, nepod, izvoz, ostalo ) "
ssql = ssql + "SELECT DISTINCTROW faktura.br_racun, faktura.br_racun AS brr, [naziv_1] & ' ' & [naziv_2] AS NAZIV, partneri.por_br, IIf([npl]<2,[vsp],0) AS iznos, IIf([npl]=0,[vsp]-[faktura].[pdv],0) AS OSN, IIf([npl]=0,[faktura].[pdv],0) AS pdv, faktura.datum, IIf([npl]=1,[faktura].[vsp]-[faktura].[pdv],0) AS kosn, IIf([npl]=1,[faktura].[pdv],0) AS kpdv, [mjesto] & ' ' & [adresa] AS adr, IIf([npl]>=2,[vsp]-[faktura].[pdv],0) AS nepod, IIf([npl]=2,[vsp]-[faktura].[pdv],0) AS izvoz, IIf([npl]=3,[vsp]-[faktura].[pdv],0) AS ostalo "
ssql = ssql + "FROM partneri INNER JOIN faktura ON partneri.sifra_kupca = faktura.sifra_kupca;"

'01.10.98 ubaèeno nepodliježe ostalo
'ssql = "INSERT INTO KIRA ( rb, br_racun, NAZIV, por_br, vsp, OSN, pdv, datum, kosn, kpdv, adresa, nepod, izvoz ) "
'ssql = ssql + "SELECT DISTINCTROW faktura.br_racun, faktura.br_racun AS brr, [naziv_1] & ' ' & [naziv_2] AS NAZIV, partneri.por_br, IIf([npl]<>2,[vsp],0) AS iznos, IIf([npl]=0,[vsp]-[faktura].[pdv],0) AS OSN, IIf([npl]=0,[faktura].[pdv],0) AS pdv, faktura.datum, IIf([npl]=1,[faktura].[vsp]-[faktura].[pdv],0) AS kosn, IIf([npl]=1,[faktura].[pdv],0) AS kpdv, [mjesto] & ' ' & [adresa] AS adr, IIf([npl]=2,[vsp]-[faktura].[pdv],0) AS nepod, IIf([npl]=2,[vsp]-[faktura].[pdv],0) AS izvoz "
'ssql = ssql + "FROM partneri INNER JOIN faktura ON partneri.sifra_kupca = faktura.sifra_kupca;"
KorigirajQuery "qKIRA", ssql

ssql = "SELECT DISTINCTROW faktura.br_racun, partneri.naziv_1, partneri.naziv_2, faktura.sifra_kupca, faktura.datum, faktura.valuta, faktura.placeno, faktura.placeno_izn, faktura.izvod_br, faktura.izvod_dat, faktura.vbp, faktura.rab, faktura.pdv, faktura.vsp, faktura.nap, faktura.tip, faktura.oznaka, faktura.oznaka "
ssql = ssql + "FROM partneri INNER JOIN faktura ON partneri.sifra_kupca = faktura.sifra_kupca "
ssql = ssql + "ORDER BY faktura.br_racun DESC;"
KorigirajQuery "faktura_pregled", ssql

ssql = "SELECT DISTINCTROW ul_faktura.br_racun, ul_faktura.sifra_dob, partneri.naziv_1, partneri.naziv_2, ul_faktura.racun, ul_faktura.datum, ul_faktura.valuta, ul_faktura.placeno, ul_faktura.placeno_izn, ul_faktura.uk_iznos, ul_faktura.izvod_br, ul_faktura.izvod_dat, ul_faktura.rabi, ul_faktura.rabp, ul_faktura.ztp, ul_faktura.zti, ul_faktura.vbp, ul_faktura.pdvi, ul_faktura.pdvp, ul_faktura.pdvO, ul_faktura.pdvNO, ul_faktura.vsp, ul_faktura.nap, ul_faktura.tip, ul_faktura.racun "
ssql = ssql + "FROM partneri INNER JOIN ul_faktura ON partneri.sifra_kupca = ul_faktura.sifra_dob "
ssql = ssql + "ORDER BY ul_faktura.br_racun DESC;"
KorigirajQuery "ul_faktura_pregled", ssql

ssql = "SELECT DISTINCTROW faktura.oznaka, [naziv_1] & ' ' & [naziv_2] AS naziv, partneri.sifra_kupca, prodroba.* "
ssql = ssql + "FROM (partneri INNER JOIN faktura ON partneri.sifra_kupca = faktura.sifra_kupca) INNER JOIN prodroba ON faktura.br_racun = prodroba.br_racun;"
KorigirajQuery "Qprodroba", ssql

mp 0
MsgBox "Korekcija završena!", 32, "Korekcija"
End Sub

'
Sub CPrihvati_Click ()
inipisi "Ispis", "Draft", chk1.Value
inipisi "Ispis", "IspisSve", chk2.Value
ispissve = chk2
draftispis = chk1
If chk1.Value Then
    Dim i%
    For i = 0 To 5
        If optprinterport(i) Then
            inipisi "Ispis", "PrinterPort", optprinterport(i).Tag
            printerport = CInt(optprinterport(i).Tag)
            Exit For
        End If
    Next i
    inipisi "Ispis", "BezKodova", optprinterport(5).Value
    BezKodova = optprinterport(5).Value
End If
inipisi "Boje", "UR", Str(pb(0).BackColor)
inipisi "Boje", "IR", Str(pb(1).BackColor)
inipisi "Boje", "tablice", Str(pb(2).BackColor)
tgbc = pb(2).BackColor
backup = txtBackup
inipisi "Baze", "BackupDir", txtBackup
Unload Me
End Sub

Sub CZaboravi_Click ()
Unload Me
End Sub

Sub dp_Click ()
frmprnodabir.Show 1
'cmd.Action = 5
End Sub

Sub Form_Load ()
CenterForm Me
chk1 = draftispis
chk2 = ispissve
Dim i
For i = 0 To 5
    If CInt(optprinterport(i).Tag) = printerport Then
        optprinterport(i) = True
        Exit For
    End If
Next i
pb(0).BackColor = CLng(inidaj("Boje", "UR", Str(12632256)))
pb(1).BackColor = CLng(inidaj("Boje", "IR", Str(8421376)))
pb(2).BackColor = CLng(inidaj("Boje", "tablice", Str(12632256)))
txtBackup = backup
End Sub

Sub KorigirajQuery (qs$, ssql$)
On Error Resume Next

Dim q As querydef

Set q = b.OpenQueryDef(qs)
If Err <> 0 Then
    Err = 0
    Set q = b.CreateQueryDef(qs, ssql)
    MsgBox "Kreiran query " & qs & "!", 32, "Korekcija"
Else
    If UCase(q.SQL) <> UCase(ssql) Then
        q.SQL = ssql
        MsgBox "Korigiran query " & qs & "!", 32, "Korekcija"
    End If
End If

q.Close
End Sub

Sub pb_Click (Index As Integer)
cmd.Color = pb(Index).BackColor
cmd.Flags = &H1&
cmd.Action = 3
pb(Index).BackColor = cmd.Color

End Sub

