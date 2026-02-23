VERSION 2.00
Begin Form pregled_robe 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pregled prodane robe"
   ClientHeight    =   4935
   ClientLeft      =   465
   ClientTop       =   2910
   ClientWidth     =   10200
   Height          =   5445
   Icon            =   PRGROBA.FRX:0000
   Left            =   405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   10200
   Top             =   2460
   Width           =   10320
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
      Height          =   3195
      HorzLines       =   2  '3D
      Layout          =   PRGROBA.FRX:0302
      LayoutIndex     =   1
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   8
      Top             =   780
      UseBookmarks    =   -1  'True
      Width           =   9735
      WrapCellPointer =   0   'False
   End
   Begin Grid g1 
      FixedCols       =   0
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   735
      HighLight       =   0   'False
      Left            =   3990
      ScrollBars      =   1  'Horizontal
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1965
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "E:\VB\97GOD\RAC\BAZE\ZEBRA_97.MDB"
      Exclusive       =   0   'False
      Height          =   285
      Left            =   60
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "prodrobagrup"
      Top             =   420
      Width           =   1155
   End
   Begin SSCheck g 
      Caption         =   "grupiraj"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Width           =   915
   End
   Begin SSCheck Ch2 
      Caption         =   "prikaži sumu"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   1290
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   1305
   End
   Begin SSCommand cmdBrZapisa 
      Caption         =   "Br &zapisa"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2670
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   900
   End
   Begin AniPushButton AniButton1 
      Height          =   330
      Left            =   3600
      Picture         =   PRGROBA.FRX:0B19
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   360
   End
   Begin SSCommand CSort 
      Caption         =   "&Sortiraj"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   900
   End
   Begin SSCommand CFilter 
      Caption         =   "&Filter"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   870
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   855
   End
   Begin SSCommand COsvjezi 
      Caption         =   "&Osvježi"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00808000&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   855
   End
End
Option Explicit
Dim whe$, ord$, sel$

Sub AniButton1_Click ()
If data1.Recordset.RecordCount = 0 Then Exit Sub
tg1.SetFocus
tg_print Me, tg1, data1, 200
End Sub

Sub CFilter_Click ()
    FiltPunoIme = 1
   Dim sn As snapshot
   Set sn = data1.Recordset.CreateSnapshot()

   unifilter3 Me, sn

   If Len(str1) > 0 Then
      whe = str1
      osvjezi
   End If
   FiltPunoIme = 0
End Sub

Sub Ch2_Click (value As Integer)
custgrid data1, tg1, g1, value
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
   Set t = d.OpenTable("prodroba")
   t.Index = "PrimaryKey"
   'Debug.Print Timer - st
   DoEvents
   mp 0
   'MsgBox "Tablica sadrži " & s.RecordCount & " zapis" & IIf(Right(rc, 1) = "1" And Right(rc, 2) <> "11", "", "a") & "!", 64, "Broj zapisa"
   MsgBox "Prikazano je " & rc & " zapis" & IIf(Right(rc, 1) = "1" And Right(rc, 2) <> "11", "", "a") & " od ukupno " & t.RecordCount & " zapisa!", 64, "Broj zapisa"
End Sub

Sub COsvjezi_Click ()
   ord = ""
   whe = ""
   osvjezi

End Sub

Sub CSort_Click ()
FiltPunoIme = 1
   Dim sn As snapshot
   Set sn = data1.Recordset.CreateSnapshot()

   unisort3 Me, sn
   If Len(str1) > 0 Then
      ord = str1
      osvjezi
   End If
FiltPunoIme = 0
End Sub

Sub Data1_Validate (action As Integer, save As Integer)
If action <> 6 Then save = False
End Sub

Sub Form_Load ()

tg1.LayoutIndex = 1
data1.DatabaseName = putbaza
data1.RecordSource = "select * from Qprodroba"
tg1.HeadBackColor = CLng(inidaj("Boje", "tablice", Str(8421376)))
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
If ch2 Then g1.Width = scalewidth - g1.Left
End Sub

Sub g_Click (value As Integer)

'ch2 = False
If value Then
'    data1.RecordSource = "select * from Prodroba"
'    tg1.LayoutIndex = 2
Else
'    tg1.LayoutIndex = 1
'    data1.RecordSource = "select * from ProdrobaGrup"
End If
osvjezi
'data1.Refresh
End Sub

Sub osvjezi ()
On Error GoTo errosvj
If Not G Then
    sel = "select * from Qprodroba"
    tg1.LayoutIndex = 1
    cmdBrZapisa.Enabled = True
Else
    sel = "select * from prodrobaGrup "
    tg1.LayoutIndex = 2
     cmdBrZapisa.Enabled = False
End If
NamjestiTGBC Me
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
If g1.Visible Then RacSumu data1, tg1, g1
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

Sub tg1_ColumnSumChange (Col As Integer)

'If tg1.LayoutIndex = 1 Then
'kol = Format(tg1.ColumnSum(4), "Standard")
'izn = Format(tg1.ColumnSum(6), "Standard")
'Else
'kol = Format(tg1.ColumnSum(3), "Standard")
'izn = Format(tg1.ColumnSum(5), "Standard")
'End If
End Sub

