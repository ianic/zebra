VERSION 2.00
Begin Form cijenik 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Cjenik"
   ClientHeight    =   5595
   ClientLeft      =   1245
   ClientTop       =   1530
   ClientWidth     =   10560
   Height          =   6105
   Icon            =   CIJENIK.FRX:0000
   Left            =   1185
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   10560
   Top             =   1080
   Width           =   10680
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "grupa"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   20
      Top             =   360
      Width           =   3555
      Begin ComboBox cboGR 
         Height          =   300
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   30
         Width           =   2985
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "csp"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   2
      Left            =   3900
      TabIndex        =   18
      Top             =   720
      Width           =   2115
      Begin MaskEdBox mask 
         DataField       =   "csp"
         DataSource      =   "Data1"
         Format          =   "#,##0.00##"
         Height          =   285
         Index           =   2
         Left            =   390
         PromptChar      =   "_"
         TabIndex        =   19
         Top             =   30
         Width           =   1665
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "PDV                    %"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   1
      Left            =   2100
      TabIndex        =   17
      Top             =   720
      Width           =   1815
      Begin MaskEdBox mask 
         DataField       =   "pdvp"
         DataSource      =   "Data1"
         Format          =   "#,##0.00##"
         Height          =   285
         Index           =   1
         Left            =   510
         PromptChar      =   "_"
         TabIndex        =   3
         Top             =   30
         Width           =   1065
      End
   End
   Begin SSCommand cmdBrZapisa 
      Caption         =   "Br &zapisa"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2700
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1110
      Width           =   900
   End
   Begin AniPushButton AniButton1 
      Height          =   330
      Left            =   3720
      Picture         =   CIJENIK.FRX:0302
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      Width           =   360
   End
   Begin SSCommand cnovi 
      Caption         =   "&Novi"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   4170
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1110
      Width           =   855
   End
   Begin SSCommand Cbrisi 
      Caption         =   "&Briši"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   5040
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1110
      Width           =   900
   End
   Begin SSCommand CSort 
      Caption         =   "&Sortiraj"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1770
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1110
      Width           =   900
   End
   Begin SSCommand CFilter 
      Caption         =   "&Filter"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   900
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1110
      Width           =   855
   End
   Begin SSCommand COsvjezi 
      Caption         =   "&Osvježi"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00808000&
      Height          =   315
      Left            =   30
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1110
      Width           =   855
   End
   Begin SSCommand Cupisi 
      Caption         =   "&Upiši"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5970
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1110
      Width           =   900
   End
   Begin SSPanel Panel3D1 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   5640
      TabIndex        =   8
      Top             =   360
      Width           =   1215
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "C:\VB\DATA\R97\ZEBRA_95.MDB"
         Exclusive       =   0   'False
         Height          =   285
         Left            =   30
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "roba"
         Top             =   30
         Width           =   1155
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
      Height          =   2685
      HorzLines       =   2  '3D
      Layout          =   CIJENIK.FRX:05F4
      LayoutIndex     =   1
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      UseBookmarks    =   -1  'True
      Width           =   9375
      WrapCellPointer =   0   'False
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "cbp"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   2115
      Begin MaskEdBox mask 
         DataField       =   "cijenarobe"
         DataSource      =   "Data1"
         Format          =   "#,##0.00##"
         Height          =   285
         Index           =   0
         Left            =   390
         PromptChar      =   "_"
         TabIndex        =   2
         Top             =   30
         Width           =   1665
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "jed. mj."
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   9
      Left            =   3540
      TabIndex        =   5
      Top             =   360
      Width           =   1395
      Begin TextBox Text1 
         DataField       =   "jed_mjere"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   600
         MaxLength       =   4
         TabIndex        =   1
         Top             =   30
         Width           =   735
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "naziv artikla"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6855
      Begin TextBox na 
         DataField       =   "nazivrobe"
         DataSource      =   "Data1"
         Height          =   300
         Left            =   960
         MaxLength       =   50
         TabIndex        =   0
         Top             =   30
         Width           =   5865
      End
   End
   Begin Label lblgrupa 
      Caption         =   "grupa"
      DataField       =   "grupa"
      DataSource      =   "Data1"
      Height          =   225
      Left            =   7290
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   1365
   End
End
Option Explicit
Dim whe$, ord$, sel$

Sub AniButton1_Click ()
If data1.Recordset.RecordCount = 0 Then Exit Sub
tg1.SetFocus
tg_print Me, tg1, data1, 100
End Sub

Sub Cbrisi_Click ()
If data1.EditMode = 2 Then
    data1.UpdateControls
    data1.Refresh
Else
    
    data1.Recordset.Delete
End If
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

Sub cnovi_Click ()
On Error GoTo cnovierr
 data1.Recordset.AddNew
 na.SetFocus
 Exit Sub
cnovierr:
    greska
    Exit Sub
End Sub

Sub COsvjezi_Click ()
   ord = ""
   whe = ""
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
lblgrupa.Tag = "NE"
lblgrupa = cbogr
lblgrupa.Tag = ""
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
Case 3022: MsgBox "Zapis s upisanin nazivom veæ postoji!", 16, "Greška"

Case Else
    MsgBox Error, 16, "Greška"
    data1.UpdateControls
End Select

response = 0

End Sub

Sub Data1_Validate (action As Integer, save As Integer)
   If action = 7 Then If MsgBox("Jeste li sigurni da želite obrisati zapis?", 36, "Provjera") = 7 Then action = 0
'   If action = 11 Then save = False

If action <> 6 Then save = False
End Sub

Sub Form_Load ()
NamjestiTGBC Me
data1.DatabaseName = putbaza
data1.RecordSource = "select * from roba"
sel = data1.RecordSource
Dim t As table
Set t = b.OpenTable("roba_grupe")
t.Index = "PrimaryKey"
Do Until t.EOF
    cbogr.AddItem ds(t!grupa)
    t.MoveNext
Loop
End Sub

Sub Form_Paint ()
form_resize
End Sub

Sub form_resize ()
On Error Resume Next
tg1.Left = 0
tg1.Width = scalewidth
tg1.Height = scaleheight - tg1.Top
End Sub

Sub lblgrupa_Change ()
On Error Resume Next
If lblgrupa.Tag = "NE" Then Exit Sub'zabrani kod upisa
If cbogr.ListCount = 0 Then Exit Sub
Dim i%
For i = 0 To cbogr.ListCount - 1
    If lblgrupa = cbogr.List(i) Then
        cbogr.ListIndex = i
        Exit For
    End If
Next i
End Sub

Sub mask_KeyPress (Index As Integer, keyascii As Integer)
keyascii = dblkp(keyascii)
If keyascii = 13 Then mask_LostFocus Index
End Sub

Sub mask_LostFocus (Index As Integer)
Dim d() As Double, i%
ReDim d(0 To 3)
On Error Resume Next
d(0) = CDbl(mask(0))
d(1) = CDbl(mask(1))
d(2) = CDbl(mask(2))

Select Case Index
Case 0, 1
    d(2) = d(0) * (1 + d(1) / 100)
Case 2
    d(0) = d(2) / (1 + d(1) / 100)
End Select

For i = 0 To 2
    mask(i) = Format(d(i), mask(i).Format)
Next i

End Sub

Sub osvjezi ()
On Error GoTo errosvj
mp 11
Dim i1%, i2%, s$
i2 = InStr(LCase(sel), "order by") - 1
If i2 = -1 Then i2 = Len(sel)
s = Left(sel, i2)

'i2 = InStr(S, "where") - 1
'If i2 = -1 Then i2 = Len(S)
'S = Left(S, i2)


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
    'Resume
    mp 0
    ord = ""
    whe = ""
    
    MsgBox Error
    data1.RecordSource = sel
    data1.Refresh
    Exit Sub
End Sub

Sub Text2_KeyPress (keyascii As Integer)
'keyascii = dblkp(keyascii)
End Sub

Sub Text2_LostFocus ()
'dbllf text2
End Sub

