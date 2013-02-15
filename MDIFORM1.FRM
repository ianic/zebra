VERSION 2.00
Begin MDIForm FRM_MDI 
   Caption         =   "Raèuni"
   ClientHeight    =   6360
   ClientLeft      =   1320
   ClientTop       =   1950
   ClientWidth     =   6045
   Height          =   7050
   Icon            =   MDIFORM1.FRX:0000
   Left            =   1260
   LinkTopic       =   "MDIForm1"
   Top             =   1320
   Width           =   6165
   WindowState     =   2  'Maximized
   Begin SSPanel Panel3D1 
      Align           =   1  'Align Top
      Font3D          =   3  'Inset w/light shading
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6045
      Begin Timer Htimer 
         Left            =   2250
         Top             =   60
      End
      Begin SSCommand cs 
         BevelWidth      =   1
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   500
         Index           =   7
         Left            =   2700
         Picture         =   MDIFORM1.FRX:0302
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   500
      End
      Begin SSCommand cs 
         BevelWidth      =   1
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   500
         Index           =   6
         Left            =   7650
         Picture         =   MDIFORM1.FRX:0604
         TabIndex        =   10
         Top             =   0
         Width           =   500
      End
      Begin SSPanel Panel3D3 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         BorderWidth     =   0
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         Height          =   375
         Left            =   10710
         TabIndex        =   9
         Top             =   60
         Width           =   1095
      End
      Begin Timer Timer2 
         Interval        =   5000
         Left            =   11520
         Top             =   60
      End
      Begin SSPanel Panel3D2 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         BorderWidth     =   0
         Caption         =   "Panel3D2"
         Font3D          =   3  'Inset w/light shading
         Height          =   375
         Left            =   8310
         TabIndex        =   8
         Top             =   60
         Width           =   2385
      End
      Begin SSPanel SL 
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         BorderWidth     =   0
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   405
         Left            =   3810
         TabIndex        =   7
         Top             =   60
         Width           =   3765
      End
      Begin SSCommand cs 
         AutoSize        =   1  'Adjust Picture Size To Button
         BevelWidth      =   1
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   500
         Index           =   0
         Left            =   0
         Picture         =   MDIFORM1.FRX:0906
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   500
      End
      Begin SSCommand cs 
         AutoSize        =   1  'Adjust Picture Size To Button
         BevelWidth      =   1
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   500
         Index           =   2
         Left            =   1680
         Picture         =   MDIFORM1.FRX:0C08
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   500
      End
      Begin SSCommand cs 
         AutoSize        =   1  'Adjust Picture Size To Button
         BevelWidth      =   1
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   500
         Index           =   1
         Left            =   660
         Picture         =   MDIFORM1.FRX:0F0A
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   500
      End
      Begin SSCommand cs 
         BevelWidth      =   1
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   500
         Index           =   5
         Left            =   3210
         Picture         =   MDIFORM1.FRX:120C
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   500
      End
      Begin SSCommand cs 
         AutoSize        =   1  'Adjust Picture Size To Button
         BevelWidth      =   1
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   500
         Index           =   4
         Left            =   1170
         Picture         =   MDIFORM1.FRX:150E
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   500
      End
      Begin SSCommand cs 
         AutoSize        =   2  'Adjust Button Size To Picture
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   600
         Index           =   3
         Left            =   12090
         Picture         =   MDIFORM1.FRX:1810
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin Menu mnu_0 
      Caption         =   "&Izlazni raèuni"
      Begin Menu mnu_1 
         Caption         =   "&Novi"
         Shortcut        =   ^N
      End
      Begin Menu mnu_2 
         Caption         =   "&Pregled"
         Shortcut        =   ^P
      End
      Begin Menu mnu_10 
         Caption         =   "Pregled &robe"
         Shortcut        =   ^R
      End
      Begin Menu mnu_3 
         Caption         =   "&Izvodi"
         Shortcut        =   ^I
      End
   End
   Begin Menu mnu_5 
      Caption         =   "&Ulazni raèuni"
      Begin Menu mnu_6 
         Caption         =   "&Novi"
      End
      Begin Menu mnu_7 
         Caption         =   "&Pregled"
      End
      Begin Menu mnu_8 
         Caption         =   "&Izvodi"
      End
   End
   Begin Menu mnu_4 
      Caption         =   "&Matièni podaci"
      Begin Menu mnu_partneri 
         Caption         =   "&Partneri"
      End
      Begin Menu mnu_cijenik 
         Caption         =   "&Cjenik"
      End
      Begin Menu mnusep01 
         Caption         =   "-"
      End
      Begin Menu mnu_GrupeRobe 
         Caption         =   "&Grupe robe"
      End
      Begin Menu mnujed_mj 
         Caption         =   "&Jedinice mjere"
      End
      Begin Menu mnuur_roba 
         Caption         =   "&Roba s ulaznih raèuna"
      End
      Begin Menu mnu_sep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin Menu mnuPostavke 
         Caption         =   "Postavke"
         Visible         =   0   'False
      End
   End
   Begin Menu mnuPregledi 
      Caption         =   "&Pregledi"
      Begin Menu mnuPdvPoRaz 
         Caption         =   "PDV po razdobljima"
      End
      Begin Menu mnuKIRA 
         Caption         =   "Knjiga &IR-a"
      End
      Begin Menu mnuKURA 
         Caption         =   "Knjiga &UR-a"
      End
   End
   Begin Menu mnuAlati 
      Caption         =   "&Alati"
      Begin Menu mnuZbrajalica 
         Caption         =   "&Zbrajalica"
      End
      Begin Menu mnusep10 
         Caption         =   "-"
      End
      Begin Menu mnu_comp 
         Caption         =   "&Popravak/kompresija baze"
      End
      Begin Menu mnubackup 
         Caption         =   "&Backup"
      End
   End
   Begin Menu mnu_proz 
      Caption         =   "P&rozori"
      Begin Menu mnu_prozori 
         Caption         =   "&Kaskada"
         Index           =   1
         Shortcut        =   +{F5}
      End
      Begin Menu mnu_prozori 
         Caption         =   "Posloži (&vodorano)"
         Index           =   2
         Shortcut        =   +{F4}
      End
      Begin Menu mnu_prozori 
         Caption         =   "Posloži (&okomito)"
         Index           =   3
      End
      Begin Menu mnu_prozori 
         Caption         =   "Posloži ikone"
         Index           =   4
      End
      Begin Menu mnu_prozori 
         Caption         =   "&Zatvori sve"
         Index           =   5
         Shortcut        =   ^{F3}
      End
      Begin Menu mnu_prozori 
         Caption         =   "-"
         Index           =   6
      End
      Begin Menu mnu_prozori 
         Caption         =   "&O programu"
         Index           =   7
      End
   End
End
Option Explicit
Dim ul_izvodi As New izvodi
Dim iz_izvodi As New izvodi
Dim gruper As New uni_ed
Dim frmjmj As New uni_ed
Dim frmkira As New knjiga
Dim frmkura As New knjiga

Sub cs_Click (Index As Integer)
Select Case Index
Case 0:
   Unload frm_mdi
Case 1: mnu_1_Click
Case 2: mnu_3_Click

Case 4: mnu_2_Click
Case 5:  mnu_partneri_Click
Case 6
   Dim X
   DoEvents
   On Error GoTo ukljuciCalc
   AppActivate "Calculator"
   Exit Sub
Case 7: mnu_cijenik_Click
End Select
Exit Sub
ukljuciCalc:
   On Error Resume Next
   X = Shell("calc.exe", 1)
   DoEvents
   Exit Sub
End Sub

Sub cs_MouseMove (Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim s$
Select Case Index
Case 0: s = "Izlaz iz programa"
Case 1: s = "Izlazni raèuni - novi"
Case 2: s = "Izlazni raèuni - izvodi"
Case 3: s = "Knjiga osnovnih sredstava"
Case 4: s = "Izlazni raèuni - pregled"
Case 5: s = "Partneri"
Case 6: s = "Kalkulator"
Case 7: s = "Cjenik"
End Select
CommandMove cs(Index), s, X, Y
End Sub

Sub cs_MouseUp (Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
capturestatus = 0
cs_MouseMove Index, 0, 0, 1, 1
End Sub

Sub Htimer_Timer ()
h ""
htimer.Interval = 0
End Sub

Sub MDIForm_Load ()

MakniItalic Me
Set b = OpenDatabase(putbaza)

'dodaj kolonu broj_narudzbenicu u faktura tablicu
Dim t As Table
Set t = b.OpenTable("faktura")
t.Index = "PrimaryKey"
Dim fc
fc = t.Fields.Count
t.Close

If fc = 22 Then
    Dim Fd As New Field
    Const DB_TEXT = 10
    Fd.Name = "broj_narudzbenice"
    Fd.Type = DB_TEXT
    Fd.Size = 100
    b.TableDefs("faktura").Fields.Append Fd
End If

Timer2_Timer
End Sub

Sub MDIForm_QueryUnload (cancel As Integer, UnloadMode As Integer)
If MsgBox("Izlaz iz programa?", 36, "Provjera") = 7 Then cancel = True
End Sub

Sub mnu_1_Click ()
mode = 0
showw iz_rac
End Sub

Sub mnu_10_Click ()
Dim b As New pregled_robe
b.Show
End Sub

Sub mnu_2_Click ()
mode = 1
Dim b As New rac_pregled
b.Show
End Sub

Sub mnu_3_Click ()
mode = 1
showw iz_izvodi
End Sub

Sub mnu_6_Click ()
mode = 0
showw ul_rac
End Sub

Sub mnu_7_Click ()
mode = 2
Dim b As New rac_pregled
b.Show
End Sub

Sub mnu_8_Click ()
mode = 2
showw ul_izvodi
End Sub

Sub mnu_cijenik_Click ()
Dim t As Table
Set t = b.OpenTable("roba")
On Error GoTo mnu_cijenik_err
showw cijenik
Exit Sub
mnu_cijenik_err:
Exit Sub
End Sub

Sub mnu_comp_Click ()
If MsgBox("Poravak i kompresija baze jeste li spremni za poèetak?", 36, "Provjera") = 7 Then Exit Sub
mp 11
On Error GoTo comp_greska
mnu_prozori_Click 5
mp 11
b.Close
sl.Caption = "Popravak baze"
RepairDatabase (putbaza)
Dim pom$
pom = Left(putbaza, Len(putbaza) - 3) & "bak"
On Error Resume Next
Kill pom
On Error GoTo comp_greska
sl.Caption = "Kompresija baze"
CompactDatabase putbaza, pom
Kill putbaza
Name pom As putbaza
Set b = OpenDatabase(putbaza)
sl.Caption = ""
Beep
mp 0
Exit Sub
comp_greska:
   mp 0
   sl.Caption = ""
   If Err = 3000 Then MsgBox "Baza mora biti zatvorena u svim drugim programima!!!", 16, "Greška" Else MsgBox Error, 16, "Greška"
   On Error Resume Next
   Set b = OpenDatabase(putbaza)
   Exit Sub
End Sub

Sub mnu_GrupeRobe_Click ()
mode = 1
str2 = "Grupe robe"
str1 = "select * from roba_grupe"
showw gruper
End Sub

Sub mnu_partneri_Click ()
showw partneri
End Sub

Sub mnu_prozori_Click (Index As Integer)
On Error Resume Next
Dim i%
Select Case Index
Case 1:
   mp 11
    frm_mdi.Arrange 0
   mp 0
Case 2:
   mp 11
    frm_mdi.Arrange 1
   mp 0
Case 5:
  screen.MousePointer = 11
  For i = 1 To forms.Count - 1
    Unload activeform
  Next i
  screen.MousePointer = 0
Case 4:
   mp 11
    frm_mdi.Arrange 3
   mp 0

'   If forms.Count > 1 Then Unload activeform
'   Exit Sub
Case 3:
   mp 11
    frm_mdi.Arrange 2
   mp 0

'   If forms.Count > 1 Then activeform.WindowState = 1
'   Exit Sub
Case 7:
   oprg.Show 1
End Select
End Sub

Sub mnubackup_Click ()
mnu_prozori_Click 5
mp 11
b.Close
DoEvents
'b.Close
h "BACKUP"

frmzip.Show 1

Set b = OpenDatabase(putbaza)
h ""
mp 0

End Sub

Sub mnujed_mj_Click ()
mode = 2
str2 = "Jedinice mjere"
str1 = "select * from jed_mj"
showw frmjmj
End Sub

Sub mnuKIRA_Click ()
mode = 2
showw frmkira
End Sub

Sub mnuKURA_Click ()
mode = 1
showw frmkura
End Sub

Sub mnuPdvPoRaz_Click ()
showw pdvor
End Sub

Sub mnuPostavke_Click ()
postavke.Show 1
End Sub

Sub mnuur_roba_Click ()
showw ul_roba
End Sub

Sub mnuZbrajalica_Click ()
showw frmKalkulator
End Sub

Sub Panel3D2_Click ()
MsgBox Str(pprinter(1, 1))
End Sub

Sub SL_Click ()
mp 0
End Sub

Sub Timer2_Timer ()
Dim s$
Select Case Weekday(Now)
Case 1: s = "nedjelja"
Case 2: s = "ponedjeljak"
Case 3: s = "utorak"
Case 4: s = "srijeda"
Case 5: s = "èetvrtak"
Case 6: s = "petak"
Case 7: s = "subota"
End Select
s = s & ", " & Format(Now, "dd.mm.yy")
If panel3d2.Caption <> s Then panel3d2.Caption = s
Panel3D3.Caption = Format(Now, "hh:nn")
End Sub

