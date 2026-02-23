VERSION 2.00
Begin Form izvodi 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Izlazni raèuni - izvodi"
   ClientHeight    =   5850
   ClientLeft      =   390
   ClientTop       =   1665
   ClientWidth     =   10590
   Height          =   6360
   Icon            =   IZVODI.FRX:0000
   Left            =   330
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   10590
   Top             =   1215
   Width           =   10710
   Begin SSCommand CPlaceno 
      Caption         =   "&Plaæeno"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   60
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   420
      Width           =   900
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "plaæeno"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   6
      Left            =   5700
      TabIndex        =   16
      Top             =   390
      Width           =   2475
      Begin MaskEdBox MaskedEdit2 
         DataField       =   "placeno_izn"
         DataSource      =   "Data1"
         Format          =   "Standard"
         Height          =   315
         Left            =   840
         PromptChar      =   "_"
         TabIndex        =   3
         Top             =   30
         Width           =   1605
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "vsp"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   13
      Left            =   5700
      TabIndex        =   14
      Top             =   60
      Width           =   2475
      Begin SSPanel uk 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   900
         TabIndex        =   15
         Top             =   30
         Width           =   1515
      End
   End
   Begin SSCommand Cupisi 
      Caption         =   "&Upiši"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   7260
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1590
      Width           =   900
   End
   Begin SSPanel Panel3D5 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   1
      Caption         =   "napomena"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   795
      Left            =   30
      TabIndex        =   11
      Top             =   780
      Width           =   8145
      Begin TextBox nap 
         DataField       =   "nap"
         DataSource      =   "Data1"
         Height          =   735
         Left            =   810
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   30
         Width           =   7275
      End
   End
   Begin SSPanel Panel3D2 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelOuter      =   1  'Inset
      BorderWidth     =   1
      Caption         =   "datum izvoda"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   1
      Left            =   1890
      TabIndex        =   10
      Top             =   60
      Width           =   2355
      Begin MaskEdBox MaskedEdit1 
         DataField       =   "izvod_dat"
         Height          =   285
         Left            =   1050
         Mask            =   "##/##/####"
         MaxLength       =   10
         PromptChar      =   "_"
         TabIndex        =   2
         Top             =   30
         Width           =   1275
      End
   End
   Begin SSPanel Panel3D2 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelOuter      =   1  'Inset
      BorderWidth     =   1
      Caption         =   "broj izvoda"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   1815
      Begin TextBox Text1 
         DataField       =   "izvod_br"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   30
         Width           =   825
      End
   End
   Begin SSCommand COsvjezi 
      Caption         =   "&Osvježi"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00808000&
      Height          =   315
      Left            =   30
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1590
      Width           =   855
   End
   Begin SSCommand CFilter 
      Caption         =   "&Filter"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   900
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1590
      Width           =   855
   End
   Begin SSCommand CSort 
      Caption         =   "&Sortiraj"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1770
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1590
      Width           =   900
   End
   Begin SSPanel Panel3D1 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   2730
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "C:\VB\DATA\R97\ZEBRA_97.MDB"
         Exclusive       =   0   'False
         Height          =   285
         Left            =   30
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "faktura_pregled"
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
      HeadingHeight   =   1.2
      Height          =   3315
      HorzLines       =   2  '3D
      Layout          =   IZVODI.FRX:0302
      LayoutIndex     =   2
      Left            =   120
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   2130
      UseBookmarks    =   -1  'True
      Width           =   9945
      WrapCellPointer =   0   'False
   End
   Begin Label Label1 
      Caption         =   "Label1"
      DataField       =   "br_racun"
      DataSource      =   "Data1"
      Height          =   225
      Left            =   8370
      TabIndex        =   13
      Top             =   390
      Visible         =   0   'False
      Width           =   1095
   End
End
Option Explicit
Dim sel$, whe$, ord$

Sub CFilter_Click ()
   Dim sn As snapshot
   Set sn = data1.Recordset.CreateSnapshot()

   unifilter3 Me, sn

   If Len(str1) > 0 Then
      whe = str1
      osvjezi
   End If

End Sub

Sub COsvjezi_Click ()
   ord = ""
   whe = ""
   osvjezi
End Sub

Sub CPlaceno_Click ()
On Error Resume Next
Dim t As table
Select Case tg1.LayoutIndex
Case 1
    Set t = b.OpenTable("faktura")
Case 2
    Set t = b.OpenTable("ul_faktura")
End Select

t.Index = "izvod_br"
t.MoveLast
If IsNull(t!izvod_br) Then text1 = 1 Else text1 = t!izvod_br
If Not IsNull(t!izvod_dat) Then
    maskededit1 = Format(t!izvod_dat, "dd.mm.yyyy")
Else
    maskededit1 = Format(Now - 2, "dd.mm.yyyy")
End If

maskededit2 = Format(data1.Recordset!vsp, "Standard")
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
If Not (IsDate(maskededit1) And IsNumeric(text1) And IsNumeric(maskededit2)) Then MsgBox "Neispravno upisani podaci!", 16, "Greška": Exit Sub
data1.Recordset.Edit
If maskededit2 = 0 Then text1 = ""
data1.Recordset!izvod_dat = maskededit1
data1.Recordset!placeno_izn = maskededit2
If data1.Recordset!placeno_izn >= data1.Recordset!vsp Then data1.Recordset!placeno = True Else data1.Recordset!placeno = False
data1.Recordset.Update
End Sub

Sub Data1_Validate (action As Integer, save As Integer)
If action <> 6 Then save = False
End Sub

Sub Form_Load ()
data1.DatabaseName = putbaza
tg1.LayoutIndex = mode
Select Case mode
Case 1
    data1.RecordSource = "select * from faktura_pregled"
    caption = "Izvodi - izlazni raèuni"
    BackColor = CLng(inidaj("Boje", "IR", Str(8421376)))
Case 2
    data1.RecordSource = "select * from ul_faktura_pregled"
    caption = "Izvodi - ulazni raèuni"
    BackColor = CLng(inidaj("Boje", "UR", Str(12632256)))
    'BackColor = &H808000
End Select
NamjestiTGBC Me

sel = data1.RecordSource
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

Sub Label1_Change ()
If data1.Recordset.RecordCount = 0 Then Exit Sub
If IsDate(data1.Recordset!izvod_dat) Then
   maskededit1 = Format(data1.Recordset!izvod_dat, "dd.mm.yyyy")
Else
   maskededit1 = "__.__.____"
End If
uk(3) = Format(data1.Recordset!vsp, "Standard")
End Sub

Sub MaskedEdit1_KeyPress (keyascii As Integer)
datkp maskededit1, keyascii
End Sub

Sub MaskedEdit1_LostFocus ()
datlf maskededit1
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

Sub Text1_LostFocus ()
intlf text1
End Sub

