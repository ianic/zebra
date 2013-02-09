VERSION 2.00
Begin Form rac_pregled 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Izlazni raèuni - pregled"
   ClientHeight    =   6180
   ClientLeft      =   570
   ClientTop       =   1395
   ClientWidth     =   11070
   Height          =   6690
   Icon            =   RACPRG.FRX:0000
   Left            =   510
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   11070
   Top             =   945
   Width           =   11190
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "E:\VB\97GOD\RAC\BAZE\ZEBRA_97.MDB"
      Exclusive       =   0   'False
      Height          =   285
      Left            =   60
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "ul_faktura_pregled"
      Top             =   390
      Width           =   1155
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
      Left            =   4050
      ScrollBars      =   1  'Horizontal
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1965
   End
   Begin TrueGrid tg2 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      BackColor       =   &H00C0C0C0&
      DataSource      =   "Data2"
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
      Height          =   1485
      HorzLines       =   2  '3D
      Layout          =   RACPRG.FRX:0302
      LayoutIndex     =   1
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   8
      Top             =   3630
      UseBookmarks    =   -1  'True
      Visible         =   0   'False
      Width           =   8205
      WrapCellPointer =   0   'False
   End
   Begin AniPushButton AniButton1 
      Height          =   330
      Left            =   2700
      Picture         =   RACPRG.FRX:099F
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   30
      Width           =   360
   End
   Begin SSCommand Cispravak 
      Caption         =   "&Ispravak"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3090
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   30
      Width           =   900
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
      Height          =   225
      Left            =   2670
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   1245
   End
   Begin SSCheck ch1 
      Caption         =   "prikaži stavke"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   225
      Left            =   1290
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   420
      Width           =   1305
   End
   Begin Data Data2 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "E:\VB\97GOD\RAC\BAZE\ZEBRA_97.MDB"
      Exclusive       =   0   'False
      Height          =   285
      Left            =   6300
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "ul_roba"
      Top             =   60
      Visible         =   0   'False
      Width           =   1155
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
      HeadingHeight   =   1.2
      Height          =   2475
      HorzLines       =   2  '3D
      Layout          =   RACPRG.FRX:0C91
      LayoutIndex     =   2
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   3
      Top             =   750
      UseBookmarks    =   -1  'True
      Width           =   10965
      WrapCellPointer =   0   'False
   End
   Begin SSCommand CSort 
      Caption         =   "&Sortiraj"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1770
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
      Left            =   900
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
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   855
   End
End
Option Explicit
Dim sel$, whe$, ord$

Sub AniButton1_Click ()
If data1.Recordset.RecordCount = 0 Then Exit Sub

Select Case tg1.LayoutIndex
Case 1'kod ulaznih dopusti ispis samo jednog raèuna
    str1 = data1.Recordset!br_racun
    ispis1.Show 1
Case 2
    mode = 1
End Select

Select Case mode
Case 0
   Select Case tg1.LayoutIndex
   Case 1
   If draftispis Then
    ispisur (data1.Recordset!br_racun)
   Else
    mp 11
    ispisracuna CLng(data1.Recordset!br_racun), 0, 1
    mp 0
   End If
   End Select
Case 1
   tg1.SetFocus
   tg_print Me, tg1, data1, 10 + 10 * tg1.LayoutIndex
End Select
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

Sub ch1_Click (value As Integer)
tg2.Visible = value
form_resize
tg1_RowChange
End Sub

Sub Ch2_Click (value As Integer)
custgrid data1, tg1, g1, value
End Sub

Sub Cispravak_Click ()
If data1.Recordset.RecordCount = 0 Then Exit Sub
Select Case tg1.LayoutIndex
Case 1
    mode = data1.Recordset!br_racun
    Dim i As New iz_rac
    i.Show
Case 2
    mode = data1.Recordset!br_racun
    Dim j As New ul_rac
    j.Show
End Select

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

Sub Form_Load ()
tg1.LayoutIndex = mode
tg2.LayoutIndex = mode
data1.DatabaseName = putbaza
Dim ssql$
Select Case mode
Case 1
    'ssql = "SELECT DISTINCTROW faktura.br_racun, partneri.naziv_1, partneri.naziv_2, faktura.sifra_kupca, faktura.datum, faktura.valuta, faktura.placeno, faktura.placeno_izn, faktura.izvod_br, faktura.izvod_dat, faktura.vbp, faktura.rab, faktura.pdv, faktura.vsp, faktura.nap, faktura.tip, faktura.oznaka, faktura.oznaka "
    'ssql = ssql + "FROM partneri INNER JOIN faktura ON partneri.sifra_kupca = faktura.sifra_kupca "
    'ssql = ssql + "ORDER BY faktura.br_racun DESC;"
    'data1.RecordSource = ssql'"select * from faktura_pregled"
    data1.RecordSource = "select * from faktura_pregled"
    caption = "Izlazni raèuni -pregled"
    BackColor = CLng(inidaj("Boje", "IR", Str(8421376)))
Case 2
    'ssql = "SELECT DISTINCTROW ul_faktura.br_racun, ul_faktura.sifra_dob, partneri.naziv_1, partneri.naziv_2, ul_faktura.racun, ul_faktura.datum, ul_faktura.valuta, ul_faktura.placeno, ul_faktura.placeno_izn, ul_faktura.uk_iznos, ul_faktura.izvod_br, ul_faktura.izvod_dat, ul_faktura.rabi, ul_faktura.rabp, ul_faktura.ztp, ul_faktura.zti, ul_faktura.vbp, ul_faktura.pdvi, ul_faktura.pdvp, ul_faktura.pdvO, ul_faktura.pdvNO, ul_faktura.vsp, ul_faktura.nap, ul_faktura.tip, ul_faktura.racun "
    'ssql = ssql + "FROM partneri INNER JOIN ul_faktura ON partneri.sifra_kupca = ul_faktura.sifra_dob "
    'ssql = ssql + "ORDER BY ul_faktura.br_racun DESC;"
    'data1.RecordSource = ssql'"select * from ul_faktura_pregled"
    data1.RecordSource = "select * from ul_faktura_pregled"
    caption = "Ulazni raèuni -pregled"
    BackColor = CLng(inidaj("Boje", "UR", Str(12632256)))
    'BackColor = &H808000
End Select
NamjestiTGBC Me
sel = data1.RecordSource
data2.DatabaseName = putbaza
data2.RecordSource = ""

End Sub

Sub Form_Paint ()
   form_resize
End Sub

Sub form_resize ()
On Error Resume Next
tg1.Left = 0
tg2.Left = 0
If ch1 Then
    tg1.Width = scalewidth
    tg2.Width = scalewidth
    'tg1.Height = 2745

    tg2.Top = scaleheight - tg2.Height'tg1.Top + tg1.Height
    tg1.Height = scaleheight - tg2.Height - tg1.Top
    'tg2.height = scaleheight - tg1.Top - tg1.height
Else
    tg1.Width = scalewidth
    tg1.Height = scaleheight - tg1.Top
End If
If ch2 Then g1.Width = scalewidth - g1.Left
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

Sub pozg1 ()
'On Error Resume Next
'Dim i%, g1w As Long
'For i = 0 To 5
'   If tg1.ColumnRectWidth(i + 5) > 0 Then g1.ColWidth(i) = tg1.ColumnRectWidth(i + 5) + 20: g1w = g1w + 3 Else g1.ColWidth(i) = 1
'Next i
''g1.ColWidth(0) = tg1.ColumnRectWidth(5) + 20
''g1.ColWidth(1) = tg1.ColumnRectWidth(6) + 20
''g1.ColWidth(2) = tg1.ColumnRectWidth(7) + 20
''g1.ColWidth(3) = tg1.ColumnRectWidth(8) + 20
''g1.ColWidth(4) = tg1.ColumnRectWidth(9) + 20
''g1.Left = tg1.ColumnRectWidth(1) + tg1.ColumnRectWidth(2) + tg1.ColumnRectWidth(3) + tg1.ColumnRectWidth(4) + 100
'g1.Left = tg1.ColumnRectLeft(5) - 10 - g1w
'g1.Width = g1.ColWidth(0) + g1.ColWidth(1) + g1.ColWidth(2) + g1.ColWidth(3) + g1.ColWidth(4) + g1.ColWidth(5) + 80

End Sub

Sub tg1_ColumnSumChange (Col As Integer)
'g1.Col = 0: g1.Text = Format(tg1.ColumnSum(5), "Standard")
'g1.Col = 1: g1.Text = Format(tg1.ColumnSum(6), "Standard")
'g1.Col = 2: g1.Text = Format(tg1.ColumnSum(7), "Standard")
'g1.Col = 4: g1.Text = Format(tg1.ColumnSum(9), "Standard")
'g1.Col = 5: g1.Text = Format(tg1.ColumnSum(10), "Standard")
End Sub

Sub tg1_RowChange ()
If Not ch1 Then Exit Sub
If data1.Recordset.RecordCount = 0 Then Exit Sub
Select Case tg1.LayoutIndex
Case 1
    data2.RecordSource = "select * from prodroba where br_racun=" & data1.Recordset!br_racun
Case 2
    data2.RecordSource = "select * from ul_roba where br_racun=" & data1.Recordset!br_racun
End Select

data2.Refresh
End Sub

