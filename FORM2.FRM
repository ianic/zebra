VERSION 2.00
Begin Form izbor 
   Caption         =   "Izbor"
   ClientHeight    =   4875
   ClientLeft      =   1035
   ClientTop       =   1350
   ClientWidth     =   9420
   Height          =   5385
   KeyPreview      =   -1  'True
   Left            =   975
   LinkTopic       =   "Form2"
   ScaleHeight     =   4875
   ScaleWidth      =   9420
   Top             =   900
   Width           =   9540
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "D:\ZEBRA\98GOD\BAZE\ZEBRA_98.MDB"
      Exclusive       =   0   'False
      Height          =   285
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "Qul_robapr"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1185
   End
   Begin TrueGrid tg1 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
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
      Height          =   3255
      HorzLines       =   0  'None
      Layout          =   FORM2.FRX:0000
      LayoutIndex     =   3
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   0
      Top             =   0
      UseBookmarks    =   -1  'True
      Width           =   9225
      WrapCellPointer =   0   'False
   End
End
Option Explicit

Sub Form_KeyPress (keyascii As Integer)
If keyascii = 13 Then tg1_DblClick
If keyascii = 27 Then Unload Me
End Sub

Sub Form_Load ()
tg1.HeadBackColor = CLng(inidaj("Boje", "tablice", Str(8421376)))
tg1.InactiveBackColor = tg1.HeadBackColor

Width = screen.Width * .95
Height = screen.Height * .9
data1.DatabaseName = putbaza
data1.RecordSource = str1
str1 = ""
tg1.LayoutIndex = mode
'If mode = 2 Then Width = 6600
CenterForm Me
End Sub

Sub form_resize ()
tg1.Width = Width - 100
tg1.Height = Height - 380
End Sub

Sub tg1_DblClick ()
If data1.Recordset.RecordCount > 0 Then
Select Case tg1.LayoutIndex
Case 2  'za izbor robe zbog PrimaryKey-a
    str1 = data1.Recordset!nazivrobe
    str2 = data1.Recordset!grupa
Case 3
    str1 = data1.Recordset!nazivrobe
    str2 = data1.Recordset!sifra_dob
    dbl1 = data1.Recordset!cijenarobe
Case Else
    str1 = data1.Recordset.Fields(0).Value
End Select
End If
Unload Me
End Sub

