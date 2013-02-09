VERSION 2.00
Begin Form ul_roba 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   4560
   ClientLeft      =   1290
   ClientTop       =   1665
   ClientWidth     =   8805
   Height          =   4965
   Icon            =   P-ULROBA.FRX:0000
   KeyPreview      =   -1  'True
   Left            =   1230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   8805
   Top             =   1320
   Width           =   8925
   Begin AniPushButton AniButton1 
      Height          =   330
      Left            =   5940
      Picture         =   P-ULROBA.FRX:0302
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   360
   End
   Begin SSCommand cmdKREIRAJ 
      Caption         =   "&KREIRAJ"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   6480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1125
   End
   Begin SSCommand CFilter 
      Caption         =   "&Filter"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3990
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   855
   End
   Begin SSCommand CSort 
      Caption         =   "&Sortiraj"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4860
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   900
   End
   Begin SSPanel Panel3D3 
      Alignment       =   0  'Left Justify - TOP
      AutoSize        =   1  'AutoSize Panel Width To Caption
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "Ins - novi     Del - briši     F5- osvježi"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   5940
      TabIndex        =   6
      Top             =   3570
      Visible         =   0   'False
      Width           =   2490
   End
   Begin SSCommand cnovi 
      Caption         =   "&Novi"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   855
   End
   Begin SSCommand Cbrisi 
      Caption         =   "&Briši"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2190
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   900
   End
   Begin SSCommand COsvjezi 
      Caption         =   "&Osvježi"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00808000&
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   855
   End
   Begin SSCommand Cupisi 
      Caption         =   "&Upiši"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7110
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3990
      Visible         =   0   'False
      Width           =   900
   End
   Begin SSPanel Panel3D1 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "D:\ZEBRA\98GOD\BAZE\ZEBRA_98.MDB"
         Exclusive       =   0   'False
         Height          =   285
         Left            =   30
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "qul_robapr"
         Top             =   30
         Width           =   1155
      End
   End
   Begin TrueGrid tg1 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      BackColor       =   &H00C0C0C0&
      DataSource      =   "Data1"
      Editable        =   -1  'True
      EditDropDown    =   -1  'True
      ExposeCellMode  =   0  'Expose upon selection
      FetchMode       =   0  'By cell
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeadingHeight   =   1
      Height          =   2685
      HorzLines       =   2  '3D
      Layout          =   P-ULROBA.FRX:05F4
      LayoutIndex     =   1
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   390
      UseBookmarks    =   -1  'True
      Width           =   8715
      WrapCellPointer =   0   'False
   End
End
Option Explicit
Dim whe$, ord$, sel$

Sub AniButton1_Click ()
If data1.Recordset.RecordCount = 0 Then Exit Sub
tg1.SetFocus
tg_print Me, tg1, data1, 1100
End Sub

Sub Cbrisi_Click ()
On Error GoTo novibris
If MsgBox("Brisanje zapisa", 20, "Brisanje") = 6 Then data1.Recordset.Delete
Exit Sub
novibris:
   If Err = 3200 Then MsgBox Error, 16, "Greška"
   data1.Recordset.MoveFirst
   tg1.SetFocus
   Exit Sub
End Sub

Sub CFilter_Click ()
   Dim sn As snapshot
   Set sn = data1.Recordset.CreateSnapshot()

   unifilter3 Me, sn

   If Len(str1) > 0 Then
      whe = str1
      osvjezi
   End If

End Sub

Sub cmdBrZapisa_Click ()
    Dim st
    st = Timer
    Debug.Print Timer
    If data1.Recordset.RecordCount = 0 Then Exit Sub
   Dim s As snapshot, rc$
   On Error Resume Next
   'H "RAÈUNAM BROJ ZAPISA"
   mp 11
   Set s = data1.Recordset.CreateSnapshot()
   If s.RecordCount > 0 Then s.MoveLast
   'Debug.Print Timer - st
   'r
   rc = Trim(Str$(s.RecordCount))

   Dim d As database, t As table
   Set d = OpenDatabase(data1.DatabaseName)
   Set t = d.OpenTable("roba")
   t.Index = "PrimaryKey"
   'Debug.Print Timer - st
   DoEvents
   mp 0
   'MsgBox "Tablica sadrži " & s.RecordCount & " zapis" & IIf(Right(rc, 1) = "1" And Right(rc, 2) <> "11", "", "a") & "!", 64, "Broj zapisa"
   MsgBox "Prikazano je " & rc & " zapis" & IIf(Right(rc, 1) = "1" And Right(rc, 2) <> "11", "", "a") & " od ukupno " & t.RecordCount & " zapisa!", 64, "Broj zapisa"
End Sub

Sub cmdKREIRAJ_Click ()
Dim t As table, r As table, uf As table, cijene%, i%
cijene = False
i = MsgBox("Korekcija podataka o veæ upisanoj robi?", 35 + 256, "Provjera")
If i = 2 Then Exit Sub
If i = 6 Then cijene = True
mp 11
h "Kreiram pregled"
On Error GoTo cmdkrerr
Set t = b.OpenTable("ul_robapr")
t.Index = "PrimaryKey"
Set uf = b.OpenTable("ul_faktura")
uf.Index = "PrimaryKey"
Set r = b.OpenTable("ul_roba")
r.Index = "prodroba"
r.MoveFirst

Do Until r.EOF
    uf.Seek "=", r!br_racun
    t.Seek "=", r!nazivrobe, uf!sifra_dob
    If t.NoMatch Then
        t.AddNew
            t!nazivrobe = r!nazivrobe
            t!jed_mjere = r!jed_mjere
            t!cijenarobe = r!cijenarobe
            t!sifra_dob = uf!sifra_dob
        t.Update
    Else
    If cijene Then
        t.Edit
            t!nazivrobe = r!nazivrobe
            t!jed_mjere = r!jed_mjere
            t!cijenarobe = r!cijenarobe
            t!sifra_dob = uf!sifra_dob
        t.Update
    End If
    End If
    r.MoveNext
Loop
h ""
mp 0
data1.Refresh
Exit Sub
cmdkrerr:
    h ""
    mp 0
    greska
    Exit Sub
End Sub

Sub cnovi_Click ()
On Error GoTo cnovierr
 data1.Recordset.AddNew

 tg1.SetFocus
Exit Sub
cnovierr:
    greska
    Exit Sub
End Sub

Sub COsvjezi_Click ()
    whe = ""
    ord = ""
    osvjezi

End Sub

Sub CSort_Click ()
   Dim sn As snapshot
   Set sn = data1.Recordset.CreateSnapshot()

   unisort3 Me, sn
   If Len(str1) > 0 Then
      ord = str1
      osvjezi
   End If

End Sub

Sub Cupisi_Click ()
On Error GoTo Cupisi_Click_err
data1.Recordset.Update
'osvjezi
Exit Sub
Cupisi_Click_err:
'data1.UpdateControls
Data1_Error 0, 0
Exit Sub
End Sub

Sub Data1_Error (dataerr As Integer, response As Integer)
If Err = 0 Then Err = dataerr

Select Case Err
'Case 3058: MsgBox "Šifra nije upisana! Ponovite upis šifre!", 16, "Greška"
Case 3022: MsgBox "Zapis s veæ postoji!", 16, "Greška"

Case Else
    MsgBox Error, 16, "Greška"
    data1.UpdateControls
End Select

response = 0

End Sub

Sub Data1_Validate (action As Integer, save As Integer)
'If action <> 6 Then save = False
End Sub

Sub Form_Activate ()
tg1.SetFocus
End Sub

Sub Form_KeyDown (keycode As Integer, Shift As Integer)
'tg1.SetFocus
Select Case keycode
'Case 13: data1.Recordset.Edit 'tg1.EditActive = True'Not tg1.EditActive
Case 45:    cnovi_Click
Case 46:    Cbrisi_Click
Case 116:     COsvjezi_Click
'Case 13:    Cupisi_Click
End Select


End Sub

Sub Form_Load ()
data1.DatabaseName = putbaza
caption = "Roba po ulaznim raèunima"

'data1.RecordSource = "SELECT DISTINCTROW ul_robaPR.nazivrobe, ul_robaPR.sifra_dob, partneri.naziv_1, partneri.naziv_2, ul_robaPR.jed_mjere, ul_robaPR.cijenarobe FROM partneri INNER JOIN ul_robaPR ON partneri.sifra_kupca = ul_robaPR.sifra_dob;"
data1.RecordSource = "SELECT * from qul_robapr"
sel = data1.RecordSource

NamjestiTGBC Me
End Sub

Sub Form_Paint ()
form_resize
End Sub

Sub form_resize ()
On Error Resume Next
tg1.Width = scalewidth
tg1.Height = scaleheight - tg1.Top
End Sub

Sub osvjezi ()
On Error GoTo errosvj
mp 11
Dim i1%, i2%, s$
i2 = InStr(LCase(sel), "order by") - 1
If i2 = -1 Then i2 = Len(sel)
s = Left(sel, i2)

i1 = InStr(LCase(sel), "where")

If InStr(s, ";") Then s = Left(s, InStr(s, ";") - 1)
data1.RecordSource = s
If i1 > 0 Then
   If Len(whe) > 1 Then data1.RecordSource = data1.RecordSource & " and " & whe
Else
   If Len(whe) > 1 Then data1.RecordSource = data1.RecordSource & " where " & whe
End If
If Len(ord) > 1 Then data1.RecordSource = data1.RecordSource & " order by " & ord

data1.Refresh
tg1.Refresh
DoEvents
mp 0
Exit Sub
errosvj:

    mp 0
    ord = ""
    whe = ""
    
    MsgBox Error
    data1.RecordSource = sel
    data1.Refresh
    Exit Sub
End Sub

