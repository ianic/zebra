VERSION 2.00
Begin Form knjiga 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Knjiga izlaznih raèuna"
   ClientHeight    =   5820
   ClientLeft      =   360
   ClientTop       =   1530
   ClientWidth     =   10260
   Height          =   6330
   Icon            =   KNJIGA.FRX:0000
   Left            =   300
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   10260
   Top             =   1080
   Width           =   10380
   Begin SSCommand cmdKREIRAJ 
      Caption         =   "&KREIRAJ"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   150
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   1125
   End
   Begin SSCommand COsvjezi 
      Caption         =   "&Osvježi"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00808000&
      Height          =   315
      Left            =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   30
      Width           =   855
   End
   Begin SSCommand CFilter 
      Caption         =   "&Filter"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1020
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   30
      Width           =   855
   End
   Begin SSCommand CSort 
      Caption         =   "&Sortiraj"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1890
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   900
   End
   Begin AniPushButton AniButton1 
      Height          =   330
      Left            =   2820
      Picture         =   KNJIGA.FRX:043A
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   30
      Width           =   360
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
      Left            =   2970
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   1305
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
      Left            =   4350
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1965
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "D:\ZEBRA\98GOD\BAZE\ZEBRA_98.MDB"
      Exclusive       =   0   'False
      Height          =   285
      Left            =   1650
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "select * from KURA"
      Top             =   420
      Width           =   1140
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
      HeadingHeight   =   1.2
      Height          =   2475
      HorzLines       =   2  '3D
      Layout          =   KNJIGA.FRX:072C
      LayoutIndex     =   2
      Left            =   30
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   780
      UseBookmarks    =   -1  'True
      Width           =   10125
      WrapCellPointer =   0   'False
   End
End
Option Explicit
Dim sel$, whe$, ord$, memode%

Sub AniButton1_Click ()
If data1.Recordset.RecordCount = 0 Then Exit Sub
tg1.SetFocus
tg_print2 Me, tg1, data1, 500 + 10 * memode
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

Sub Ch2_Click (value As Integer)
custgrid data1, tg1, g1, value
form_resize
End Sub

Sub cmdKREIRAJ_Click ()
If MsgBox("Kreiranje knjige", 36, "Provjera") = 7 Then Exit Sub
h "Kreiranje knjige"
mp 11
Dim ssql$, qd As querydef
On Error GoTo kreirajgreska
'Dim t As table, rb&

Select Case memode
Case 1
    Set qd = b.OpenQueryDef("qkura")
   'ssql = "INSERT INTO KURA ( rb, br_racun, racun, NAZIV, por_br, vbp, vsp, pdvi, pdvO, pdvNO ,datum ) "
   'ssql = ssql + "SELECT DISTINCTROW ul_faktura.br_racun, ul_faktura.br_racun, ul_faktura.racun, [naziv_1] & ' ' & [naziv_2] AS NAZIV, partneri.por_br, ul_faktura.vbp, ul_faktura.vsp, ul_faktura.pdvi, ul_faktura.pdvO, ul_faktura.pdvNO ,ul_faktura.datum "
   'ssql = ssql + "FROM partneri INNER JOIN ul_faktura ON partneri.sifra_kupca = ul_faktura.sifra_dob;"
   b.Execute "delete * from KURA"
   qd.Execute
   'b.Execute ssql
'   Set t = b.OpenTable("kUra")
'   t.Index = "datum"
Case 2
    Set qd = b.OpenQueryDef("qkira")
   'ssql = "INSERT INTO KIRA (rb,  br_racun, NAZIV, por_br, vsp, OSN, pdv, datum ) "
   'ssql = ssql + "SELECT DISTINCTROW faktura.br_racun, faktura.br_racun, [naziv_1] & ' ' & [naziv_2] AS NAZIV, partneri.por_br, faktura.vsp, [vsp]-[pdv] AS OSN, faktura.pdv, faktura.datum FROM partneri "
   'ssql = ssql + "INNER JOIN faktura ON partneri.sifra_kupca = faktura.sifra_kupca;"
   b.Execute "delete * from KIRA"
   qd.Execute
   'b.Execute ssql
'   Set t = b.OpenTable("kira")
'   t.Index = "PrimaryKey"
End Select


't.MoveFirst
'rb = 1
'Do Until t.EOF
'   t.Edit
'      t!rb = rb
'   t.Update
'   rb = rb + 1
'   t.MoveNext
'Loop
data1.Refresh
tg1.Refresh
mp 0
h ""
Exit Sub
kreirajgreska:
   greska
   mp 0
   h ""
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

Sub Form_Load ()
data1.DatabaseName = putbaza
Select Case mode
Case 1
   data1.RecordSource = "SELECT * FROM KURA"
   caption = "Knjiga ulaznih raèuna"
   BackColor = CLng(inidaj("Boje", "UR", Str(12632256)))
Case 2
   data1.RecordSource = "SELECT * FROM KIRA"
   caption = "Knjiga izlaznih raèuna"
   BackColor = CLng(inidaj("Boje", "IR", Str(8421376)))
End Select
memode = mode
tg1.LayoutIndex = mode
NamjestiTGBC Me
sel = data1.RecordSource
End Sub

Sub form_resize ()
On Error Resume Next
tg1.Left = 0
tg1.Width = scalewidth
tg1.Height = scaleheight - tg1.Top
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

