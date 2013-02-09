VERSION 2.00
Begin Form partneri 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Partneri"
   ClientHeight    =   6660
   ClientLeft      =   1140
   ClientTop       =   945
   ClientWidth     =   10620
   FillColor       =   &H000000C0&
   Height          =   7065
   Icon            =   MP-PAR.FRX:0000
   KeyPreview      =   -1  'True
   Left            =   1080
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   10620
   Top             =   600
   Width           =   10740
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "mjesto"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   7
      Left            =   4830
      TabIndex        =   35
      Top             =   900
      Width           =   3375
      Begin TextBox t 
         DataField       =   "mjesto"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   8
         Left            =   30
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   210
         Width           =   3315
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "porezni broj (MB ili JMBG)"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   6
      Left            =   60
      TabIndex        =   34
      Top             =   1980
      Width           =   2955
      Begin TextBox t 
         DataField       =   "por_br"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   9
         Left            =   30
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   210
         Width           =   2895
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "žiro"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   5
      Left            =   3030
      TabIndex        =   33
      Top             =   1980
      Width           =   5625
      Begin TextBox t 
         DataField       =   "ziro"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   5
         Left            =   30
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   210
         Width           =   5535
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "kontakt"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   9
      Left            =   60
      TabIndex        =   32
      Top             =   2520
      Width           =   3375
      Begin TextBox t 
         DataField       =   "osoba"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   6
         Left            =   30
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   210
         Width           =   3315
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   3450
      TabIndex        =   31
      Top             =   2520
      Width           =   1965
      Begin SSCheck Check3D1 
         Caption         =   "dobavljaè"
         DataField       =   "Dob"
         DataSource      =   "Data1"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   12
         Top             =   30
         Value           =   -1  'True
         Width           =   975
      End
      Begin SSCheck Check3D1 
         Caption         =   "kupac"
         DataField       =   "Kup"
         DataSource      =   "Data1"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   30
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "fax"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   4
      Left            =   5760
      TabIndex        =   30
      Top             =   1440
      Width           =   2895
      Begin TextBox t 
         DataField       =   "telefax"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   4
         Left            =   30
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   210
         Width           =   2805
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "telefon"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   3
      Left            =   60
      TabIndex        =   29
      Top             =   1440
      Width           =   5685
      Begin TextBox t 
         DataField       =   "telefon"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   30
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   210
         Width           =   5595
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "adresa"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   2
      Left            =   60
      TabIndex        =   28
      Top             =   900
      Width           =   4755
      Begin TextBox t 
         DataField       =   "adresa"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   210
         Width           =   4725
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "naziv"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   825
      Index           =   1
      Left            =   1530
      TabIndex        =   27
      Top             =   60
      Width           =   6225
      Begin TextBox t 
         DataField       =   "naziv_2"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   7
         Left            =   30
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   510
         Width           =   6165
      End
      Begin TextBox t 
         DataField       =   "naziv_1"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   30
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   210
         Width           =   6165
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "šifra"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   0
      Left            =   60
      TabIndex        =   26
      Top             =   60
      Width           =   1455
      Begin TextBox t 
         DataField       =   "sifra_kupca"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   30
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   210
         Width           =   1365
      End
   End
   Begin TrueGrid tg1 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      BackColor       =   &H00C0C0C0&
      DataSource      =   "Data1"
      Editable        =   0   'False
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
      Height          =   2835
      HorzLines       =   2  '3D
      Layout          =   MP-PAR.FRX:0302
      LayoutIndex     =   1
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   3060
      UseBookmarks    =   -1  'True
      Width           =   10485
      WrapCellPointer =   0   'False
   End
   Begin SSPanel Panel3D1 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BorderWidth     =   2
      Font3D          =   1  'Raised w/light shading
      Height          =   765
      Left            =   0
      ShadowColor     =   1  'Black
      TabIndex        =   13
      Top             =   5910
      Width           =   5295
      Begin AniPushButton AniButton1 
         Height          =   330
         Left            =   3630
         Picture         =   MP-PAR.FRX:0A6E
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   90
         Width           =   360
      End
      Begin SSCommand COsvjezi 
         Caption         =   "&Osvježi"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   930
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   60
         Width           =   855
      End
      Begin SSCommand cmdBrZapisa 
         Caption         =   "Br &zapisa"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2670
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   390
         Width           =   900
      End
      Begin SSPanel Panel3D2 
         Alignment       =   0  'Left Justify - TOP
         BevelOuter      =   1  'Inset
         BorderWidth     =   2
         Font3D          =   1  'Raised w/light shading
         Height          =   435
         Left            =   4050
         TabIndex        =   22
         Top             =   60
         Width           =   1185
         Begin AniPushButton AniButton2 
            Height          =   330
            Left            =   780
            PictDrawMode    =   1  'Autosize control
            Picture         =   MP-PAR.FRX:0D60
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   30
            Width           =   360
         End
         Begin SpinButton Spin1 
            BackColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   60
            LightColor      =   &H00E0FFFF&
            ShadeColor      =   &H00C0C0C0&
            ShadowForeColor =   &H00404040&
            ShadowThickness =   2
            SpinOrientation =   1  'Horizontal
            Top             =   30
            Width           =   675
         End
      End
      Begin SSCommand cupisi 
         Caption         =   "&Upiši"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   930
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   390
         Width           =   855
      End
      Begin SSCommand CFilter 
         Caption         =   "&Filter"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   60
         Width           =   855
      End
      Begin SSCommand CSort 
         Caption         =   "&Sortiraj"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2670
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   60
         Width           =   900
      End
      Begin SSCommand cbrisi 
         Caption         =   "&Briši"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   390
         Width           =   855
      End
      Begin SSCommand cnovi 
         Caption         =   "&Novi"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   60
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   390
         Width           =   855
      End
      Begin SSCommand cizaberi 
         Caption         =   "&Izaberi"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   855
      End
      Begin SSCommand cedit 
         Caption         =   "&Edit"
         Font3D          =   3  'Inset w/light shading
         Height          =   315
         Left            =   60
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin Data Data1 
      Connect         =   ""
      DatabaseName    =   "C:\VB\DATA\R97\ZEBRA_97.MDB"
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7470
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "partneri"
      Top             =   2730
      Width           =   1140
   End
End
Option Explicit
Dim whe$, ord$, sel$, memode%

Sub AniButton1_Click ()
tg1.SetFocus
tg_print Me, tg1, data1, 110
End Sub

Sub AniButton2_Click ()
tg1.SetFocus
If tg1.FontBold Then tg1.FontBold = False Else tg1.FontBold = True
End Sub

Sub Cbrisi_Click ()
If data1.Recordset.RecordCount = 0 Then Exit Sub
On Error GoTo cbrisierr
'u toku je upis novog
If data1.EditMode = 2 Then
   data1.UpdateControls
   data1.Refresh
   Exit Sub
End If

'If tg1.DataChanged Then data1.UpdateControls 'u toku je editiranje postojeæeg
data1.Recordset.Delete
Exit Sub

cbrisierr:
   data1.UpdateControls
   data1.Refresh
   Exit Sub
End Sub

Sub cbrisi_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
tg1.SetFocus
End Sub

Sub cedit_Click ()
If data1.Recordset.RecordCount = 0 Then Exit Sub
'tg1.EditActive = True
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

Sub CFilter_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
tg1.SetFocus
End Sub

Sub cizaberi_click ()
   
'   If data1.Recordset.RecordCount > 0 Then
'      cizaberi.Tag = "ok"
'      'Select Case memode
'      '      'Case 1
'            '   str1 = data1.Recordset.Fields(0).Name & "=" & data1.Recordset.Fields(0).Value
'            '   Unload Me
'      '      Case Else
'               If data1.Recordset.Fields(0).Type = 10 Then
'                  str1 = data1.Recordset.Fields(0).Name & "='" & data1.Recordset.Fields(0).Value & "'"
'               Else
'                  str1 = data1.Recordset.Fields(0).Name & "=" & data1.Recordset.Fields(0).Value
'               End If
'
'               If data1.Recordset.Fields(1).Type = 10 Then
'                  str1 = str1 & " and " & data1.Recordset.Fields(1).Name & "='" & data1.Recordset.Fields(1).Value & "'"
'               Else
'                  str1 = str1 & " and " & data1.Recordset.Fields(1).Name & "=" & data1.Recordset.Fields(1).Value
'               End If
'               Unload Me
'      'End Select
'   End If
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
   Set t = d.OpenTable("partneri")
   't.Index = "PrimaryKey"
   't.MoveFirst
   'Dim i%
   'Do Until t.EOF
   ' i = i + 1
   ' t.MoveNext
   'Loop
   't.MoveLast
   'Debug.Print Timer - st
   DoEvents
   mp 0
   'MsgBox "Tablica sadrži " & s.RecordCount & " zapis" & IIf(Right(rc, 1) = "1" And Right(rc, 2) <> "11", "", "a") & "!", 64, "Broj zapisa"
   'MsgBox "Prikazano je " & rc & " zapis" & IIf(Right(rc, 1) = "1" And Right(rc, 2) <> "11", "", "a") & " od ukupno " & t.RecordCount & " zapisa!", 64, "Broj zapisa"
   MsgBox "Prikazano je " & rc & " zapis" & IIf(Right(rc, 1) = "1" And Right(rc, 2) <> "11", "", "a") & "!", 64, "Broj zapisa"
End Sub

Sub cmdBrZapisa_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
tg1.SetFocus
End Sub

Sub cnovi_Click ()
On Error GoTo cnovierr
If data1.EditMode = 2 Then Exit Sub
   mp 11
   Dim psif$
   On Error Resume Next
   psif = InputBox("prva dva broja:", "Poštanski broj...", "32")
   On Error GoTo cnovierr
    Dim s As snapshot
    Set s = b.CreateSnapshot("select sifra_kupca from Partneri where sifra_kupca like '" & psif & "*' order by sifra_kupca")
    If s.RecordCount > 0 Then
        s.MoveLast
        If IsNumeric(psif) Then
            psif = Trim(Str(CLng(Val(s!sifra_kupca) + 1)))
        Else
            'psif = psif & Trim(Str(CLng(Val(s!sifra_kupca) + 1)))
       End If
    Else
        psif = psif + "001"
    End If
   data1.Recordset.AddNew
   t(0).Text = psif
   On Error Resume Next
   t(0).SetFocus
   mp 0
Exit Sub
cnovierr:
    greska
    mp 0
    Exit Sub
End Sub

Sub COsvjezi_Click ()
   ord = ""
   whe = ""
   osvjezi
End Sub

Sub COsvjezi_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
   tg1.SetFocus
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

Sub CSort_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
   tg1.SetFocus
End Sub

Sub Cupisi_Click ()
    mp 11
   'If Not tg1.DataChanged Then Exit Sub
   On Error GoTo cupisierr
   
   If data1.EditMode = 2 Then
      data1.Recordset.Update
   Else
      data1.UpdateRecord
   End If
   mp 0
   Exit Sub

cupisierr:
      mp 0
      Data1_Error 0, 0
      Exit Sub
End Sub

Sub cupisi_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
   tg1.SetFocus
End Sub

Sub Data1_Error (dataerr As Integer, response As Integer)
If Err = 0 Then Err = dataerr

Select Case Err
Case 3058: MsgBox "Šifra nije upisana! Ponovite upis šifre!", 16, "Greška"
Case 3022: MsgBox "Zapis s upisanom šifrom veæ postoji! Ponovite upis šifre!", 16, "Greška"

Case Else
    MsgBox Error, 16, "Greška"
    data1.UpdateControls
End Select

response = 0
End Sub

Sub Data1_Validate (action As Integer, save As Integer)
   If action = 7 Then If MsgBox("Jeste li sigurni da želite obrisati zapis?", 36, "Provjera") = 7 Then action = 0
   If action = 11 Then save = False
End Sub

Sub Form_KeyPress (KeyAscii As Integer)
On Error Resume Next
If cizaberi.Visible And KeyAscii = 13 Then
   cizaberi_click
   Exit Sub
End If
'If keyascii = 27 And tg1.EditActive = False And Me.MDIChild = False Then Unload Me
End Sub

Sub Form_Load ()
mp 11
'pregled Me, mode
NamjestiTGBC Me

data1.DatabaseName = PutBaza
data1.RecordSource = "select * from partneri"
sel = data1.RecordSource
ord = " sifra_kupca"
data1.Refresh

Dim i%

For i = 1 To 9
    t(i).MaxLength = data1.Recordset.Fields(t(i).DataField).Size
Next i

'tg_combo1 Me, 6, "_artGR1"
'tg_combo1 Me, 7, "_artGR2"

cizaberi.Tag = ""
whe = ""
'ord = ""
tg1.Refresh
If data1.Recordset.RecordCount > 0 Then data1.Recordset.MoveFirst
DoEvents
mp 0

End Sub

Sub form_resize ()
On Error Resume Next
tg1.Width = Width - 100
tg1.Height = Height - 1150 - tg1.Top

panel3d1.Top = tg1.Height + tg1.Top
panel3d1.Left = 0
End Sub

Sub Form_Unload (Cancel As Integer)
   On Error Resume Next
   If cizaberi.Tag = "" Then str1 = ""
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

Sub Spin1_SpinDown ()
   tg1.FontSize = tg1.FontSize - 3
   tg1.FontItalic = False
End Sub

Sub Spin1_SpinUp ()
   tg1.FontSize = tg1.FontSize + 3
   tg1.FontItalic = False
End Sub

