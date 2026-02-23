VERSION 2.00
Begin Form pdvor 
   BackColor       =   &H00C0C0C0&
   Caption         =   "PDV po obraèunskim razdobljima"
   ClientHeight    =   5820
   ClientLeft      =   855
   ClientTop       =   1680
   ClientWidth     =   7140
   Height          =   6225
   Icon            =   PDVOR.FRX:0000
   Left            =   795
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   7140
   Top             =   1335
   Width           =   7260
   Begin SSCommand cmdKREIRAJ 
      Caption         =   "&KREIRAJ"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   30
      Width           =   1125
   End
   Begin AniPushButton AniButton1 
      Height          =   330
      Left            =   1800
      Picture         =   PDVOR.FRX:043A
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   360
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
      Left            =   2430
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1965
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "E:\VB\97GOD\RAC\BAZE\ZEBRA_98.MDB"
      Exclusive       =   0   'False
      Height          =   285
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "PDVR2"
      Top             =   330
      Visible         =   0   'False
      Width           =   1185
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
      Height          =   5265
      HorzLines       =   2  '3D
      Layout          =   PDVOR.FRX:072C
      LayoutIndex     =   1
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   750
      UseBookmarks    =   -1  'True
      Width           =   10845
      WrapCellPointer =   0   'False
   End
End
Option Explicit

Sub AniButton1_Click ()
If data1.Recordset.RecordCount = 0 Then Exit Sub
tg1.SetFocus
tg_print Me, tg1, data1, 300

End Sub

Sub cmdKREIRAJ_Click ()
If MsgBox("Kreiranje izvještaja", 36, "Provjera") = 7 Then Exit Sub
On Error GoTo cmdkrerr
mp 11
Dim t As Table, p As Table, p1#
Set t = b.OpenTable("faktura")
Set p = b.OpenTable("obr_raz")
t.Index = "Primarykey"
p.Index = "Primarykey"

h "BRISANJE"
p.MoveFirst
Do Until p.EOF
    p.Edit
    p!p1 = 0
    p!p2 = 0
    p!pdv = 0
    p.Update
    p.MoveNext
Loop

h "IZLAZNI RAÈUNI"
t.MoveFirst
Do Until t.EOF
    If db(t!placeno_izn) > 0 Then
    If t!placeno = False Then
        If db(t!vsp) > 0 Then
            p1 = db(t!placeno_izn) * db(t!pdv) / db(t!vsp)
        Else
            p1 = 0
        End If
    Else
        p1 = db(t!pdv)
    End If
    upisipdv p, p1, t!izvod_dat, "p1"
    End If
    t.MoveNext
Loop
h "ULAZNI RAÈUNI"
Set t = b.OpenTable("ul_faktura")
t.Index = "Primarykey"
t.MoveFirst
Do Until t.EOF
    If db(t!placeno_izn) > 0 Then
    If t!placeno = False Then
        If db(t!vsp) > 0 Then
            p1 = db(t!placeno_izn) * db(t!pdvO) / db(t!vsp)
        Else
            p1 = 0
        End If
    Else
        p1 = db(t!pdvO)
    End If
    upisipdv p, p1, t!izvod_dat, "p2"
    End If
    t.MoveNext
Loop
h "SUMA"
p.MoveFirst
Do Until p.EOF
    p.Edit
    p!pdv = db(p!p1) - db(p!p2)
    p.Update
    p.MoveNext
Loop
h ""
mp 0
data1.Refresh
RacSumu data1, tg1, g1
Exit Sub
cmdkrerr:
    greska
    mp 0
    h ""
    Exit Sub
End Sub

Sub Form_Load ()
mp 11
data1.DatabaseName = putbaza
'data1.RecordSource = "select * from pdvr2"
data1.RecordSource = "select * from obr_raz where pdv<>0"
data1.Refresh
tg1.Refresh
custgrid data1, tg1, g1, True
RacSumu data1, tg1, g1
DoEvents
mp 0
NamjestiTGBC Me
End Sub

Sub form_resize ()
On Error Resume Next
tg1.Width = scalewidth
tg1.Height = scaleheight - tg1.Top
g1.Width = scalewidth - g1.Left
End Sub

Sub tg1_ExitEdit ()
RacSumu data1, tg1, g1
End Sub

Sub upisipdv (p As Table, por#, dat, polje$)
p.MoveFirst
Do Until p.EOF
    If p!datp <= dat And p!datk >= dat Then
    p.Edit
        p.Fields(polje) = db(p.Fields(polje)) + por
    p.Update
    Exit Do
    End If
    p.MoveNext
Loop
End Sub

