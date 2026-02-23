VERSION 2.00
Begin Form ul_rac 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Ulazni raèuni - novi"
   ClientHeight    =   7425
   ClientLeft      =   3300
   ClientTop       =   4065
   ClientWidth     =   9375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   7830
   Left            =   3240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   9375
   Top             =   3720
   Visible         =   0   'False
   Width           =   9495
   Begin ComboBox cboTR 
      Height          =   315
      Left            =   6540
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   945
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "vbp"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   15
      Left            =   6720
      TabIndex        =   52
      Top             =   5610
      Width           =   2475
      Begin SSPanel uk 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         DataField       =   "vbp"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   1
         Left            =   870
         TabIndex        =   53
         Top             =   30
         Width           =   1545
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "raèun"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   14
      Left            =   5610
      TabIndex        =   51
      Top             =   330
      Width           =   3735
      Begin TextBox racun 
         Height          =   300
         Left            =   480
         MaxLength       =   50
         TabIndex        =   2
         Top             =   30
         Width           =   3195
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "+"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   705
      Index           =   12
      Left            =   1770
      TabIndex        =   49
      Top             =   5250
      Width           =   2325
      Begin TextBox txt 
         DataField       =   "zti"
         Height          =   300
         Index           =   3
         Left            =   180
         MaxLength       =   18
         TabIndex        =   13
         Top             =   360
         Width           =   1545
      End
      Begin SSPanel Panel3D4 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         Caption         =   "%  zavisni troškovi"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   2
         Left            =   900
         TabIndex        =   50
         Top             =   60
         Width           =   1365
      End
      Begin TextBox txt 
         DataField       =   "ztp"
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   30
         Width           =   615
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "+"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   705
      Index           =   4
      Left            =   30
      TabIndex        =   47
      Top             =   5940
      Width           =   6675
      Begin TextBox txt 
         DataField       =   "pdvi"
         Height          =   300
         Index           =   7
         Left            =   180
         MaxLength       =   18
         TabIndex        =   15
         Top             =   360
         Width           =   1545
      End
      Begin SSPanel Panel3D4 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         Caption         =   "ne može se odbiti"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   4
         Left            =   4680
         TabIndex        =   56
         Top             =   60
         Width           =   1545
      End
      Begin SSPanel Panel3D4 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         Caption         =   "za odbiti"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   3
         Left            =   3090
         TabIndex        =   55
         Top             =   60
         Width           =   1005
      End
      Begin TextBox txt 
         DataField       =   "pdvno"
         Height          =   300
         Index           =   6
         Left            =   4680
         MaxLength       =   18
         TabIndex        =   17
         Top             =   360
         Width           =   1545
      End
      Begin SSPanel uk 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelInner      =   1  'Inset
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         DataField       =   "pdvi"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   2
         Left            =   1560
         TabIndex        =   54
         Top             =   30
         Width           =   1365
      End
      Begin TextBox txt 
         DataField       =   "pdvo"
         Height          =   300
         Index           =   5
         Left            =   3090
         MaxLength       =   18
         TabIndex        =   16
         Top             =   360
         Width           =   1545
      End
      Begin TextBox txt 
         DataField       =   "pdvp"
         Height          =   315
         Index           =   4
         Left            =   180
         TabIndex        =   14
         Text            =   "25"
         Top             =   0
         Width           =   615
      End
      Begin SSPanel Panel3D4 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         Caption         =   "%  pdv ="
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   1
         Left            =   870
         TabIndex        =   48
         Top             =   60
         Width           =   675
      End
   End
   Begin SSPanel Panel3D5 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
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
      TabIndex        =   46
      Top             =   6630
      Width           =   8325
      Begin TextBox nap 
         Height          =   735
         Left            =   810
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   30
         Width           =   7455
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "vsp"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   13
      Left            =   6690
      TabIndex        =   44
      Top             =   6270
      Width           =   2475
      Begin SSPanel uk 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         DataField       =   "vsp"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   3
         Left            =   870
         TabIndex        =   45
         Top             =   30
         Width           =   1545
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "- "
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   705
      Index           =   11
      Left            =   30
      TabIndex        =   42
      Top             =   5250
      Width           =   1755
      Begin TextBox txt 
         DataField       =   "rabi"
         Height          =   300
         Index           =   1
         Left            =   150
         MaxLength       =   18
         TabIndex        =   11
         Top             =   360
         Width           =   1545
      End
      Begin SSPanel Panel3D4 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         Caption         =   "%  rabat"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   0
         Left            =   840
         TabIndex        =   43
         Top             =   60
         Width           =   675
      End
      Begin TextBox txt 
         DataField       =   "rabp"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   30
         Width           =   615
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "S"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Symbol"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   10
      Left            =   7080
      TabIndex        =   40
      Top             =   5280
      Width           =   2115
      Begin SSPanel uk 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         DataField       =   "uk_iznos"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   0
         Left            =   480
         TabIndex        =   41
         Top             =   30
         Width           =   1545
      End
   End
   Begin SSCommand CMD1 
      Caption         =   "U&piši"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   0
      Left            =   8370
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   6690
      Width           =   975
   End
   Begin SSCommand CMD1 
      Caption         =   "&Odustani"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   1
      Left            =   8370
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   7050
      Width           =   975
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "C:\VB\DATA\R9701\BAZE\ZEBRA_95.MDB"
      Exclusive       =   0   'False
      Height          =   285
      Left            =   4710
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "ul_roba"
      Top             =   5460
      Visible         =   0   'False
      Width           =   1185
   End
   Begin SSCommand CMD 
      Caption         =   "&Novi artikl"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   6990
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2295
   End
   Begin SSCommand CMD 
      Caption         =   "&Izmjeni artikl"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2295
   End
   Begin SSCommand CMD 
      Caption         =   "&Briši artikl"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   2370
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2295
   End
   Begin SSCommand CMD 
      Caption         =   "&Upiši artikl"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2295
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
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   9
      Left            =   30
      TabIndex        =   33
      Top             =   1890
      Width           =   1755
      Begin ComboBox jm 
         Height          =   300
         Left            =   600
         TabIndex        =   7
         Top             =   30
         Width           =   1125
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "iznos"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   8
      Left            =   6090
      TabIndex        =   31
      Top             =   1890
      Width           =   2115
      Begin SSPanel izn 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   480
         TabIndex        =   32
         Top             =   30
         Width           =   1545
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "cijena"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   7
      Left            =   3990
      TabIndex        =   30
      Top             =   1890
      Width           =   2115
      Begin TextBox Text3 
         Height          =   300
         Index           =   1
         Left            =   540
         MaxLength       =   18
         TabIndex        =   9
         Top             =   30
         Width           =   1545
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "kolièina"
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
      Left            =   1770
      TabIndex        =   29
      Top             =   1890
      Width           =   2235
      Begin TextBox Text3 
         Height          =   300
         Index           =   0
         Left            =   660
         MaxLength       =   18
         TabIndex        =   8
         Top             =   30
         Width           =   1545
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
      ForeColor       =   &H00000080&
      Height          =   555
      Index           =   5
      Left            =   30
      TabIndex        =   28
      Top             =   1350
      Width           =   9315
      Begin SSPanel Panel3D2 
         Alignment       =   0  'Left Justify - TOP
         BevelOuter      =   0  'None
         BorderWidth     =   1
         Caption         =   "F3-naziv F4-naziv+dobavljaè"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   1
         Left            =   7800
         TabIndex        =   57
         Top             =   60
         Width           =   1455
      End
      Begin TextBox na 
         Height          =   300
         Left            =   930
         MaxLength       =   50
         TabIndex        =   6
         Top             =   120
         Width           =   6825
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "rok plaæanja"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   345
      Index           =   2
      Left            =   5610
      TabIndex        =   26
      Top             =   1020
      Width           =   3735
      Begin TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   930
         TabIndex        =   4
         Top             =   30
         Width           =   585
      End
      Begin MaskEdBox MaskedEdit1 
         Height          =   285
         Index           =   1
         Left            =   2460
         Mask            =   "##/##/####"
         MaxLength       =   10
         PromptChar      =   "_"
         TabIndex        =   5
         Top             =   30
         Width           =   1215
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelOuter      =   0  'None
         BorderWidth     =   1
         Caption         =   "dana, valuta"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   3
         Left            =   1530
         TabIndex        =   27
         Top             =   30
         Width           =   945
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "datum"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   345
      Index           =   1
      Left            =   7530
      TabIndex        =   25
      Top             =   690
      Width           =   1815
      Begin MaskEdBox MaskedEdit1 
         Height          =   285
         Index           =   0
         Left            =   540
         Mask            =   "##/##/####"
         MaxLength       =   10
         PromptChar      =   "_"
         TabIndex        =   3
         Top             =   30
         Width           =   1215
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "broj"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Index           =   0
      Left            =   7530
      TabIndex        =   23
      Top             =   30
      Width           =   1815
      Begin SSPanel br 
         BevelOuter      =   0  'None
         BorderWidth     =   1
         Caption         =   "12345"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   690
         TabIndex        =   24
         Top             =   60
         Width           =   1065
      End
   End
   Begin SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "dobavljaè"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   345
      Left            =   30
      TabIndex        =   21
      Top             =   30
      Width           =   5565
      Begin SSPanel Panel3D2 
         Alignment       =   0  'Left Justify - TOP
         BevelOuter      =   0  'None
         BorderWidth     =   1
         Caption         =   "F3 - šifra   F4 - naziv_1   F5 - naziv_2"
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
         Left            =   2160
         TabIndex        =   22
         Top             =   60
         Width           =   2835
      End
      Begin TextBox sk 
         Height          =   285
         Left            =   780
         MaxLength       =   11
         TabIndex        =   0
         Top             =   30
         Width           =   1305
      End
   End
   Begin SSPanel p 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   945
      Left            =   30
      TabIndex        =   20
      Top             =   390
      Width           =   5565
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
      HeadingHeight   =   1
      Height          =   2655
      HorzLines       =   2  '3D
      Layout          =   ULRAC.FRX:0000
      LayoutIndex     =   1
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2610
      UseBookmarks    =   -1  'True
      Width           =   9315
      WrapCellPointer =   0   'False
   End
   Begin Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4080
      TabIndex        =   58
      Top             =   3480
      Width           =   1215
   End
End
Option Explicit
Dim sifk$, memode%, unl%
Dim vrati%, book
Dim pdvchg%'da li je upisivano u polje PDV
            ' ako je zadrži taj pdv

Sub brisiart ()
   na = ""
   jm = ""
   text3(0) = ""
   text3(1) = ""
   izn = ""
End Sub

Sub CMD_Click (Index As Integer)
On Error GoTo cmd_err1
Select Case Index
Case 0
   Text3_LostFocus 1
   If Not (Len(na) > 0 And IsNumeric(izn) And Len(jm.Text) > 0) Then MsgBox "Neispravno upisani podaci!", 16, "Greška": Exit Sub
   If vrati Then data1.Recordset.Bookmark = book: data1.Recordset.Edit  Else data1.Recordset.AddNew
      data1.Recordset!br_racun = br
      data1.Recordset!nazivrobe = na
      data1.Recordset!jed_mjere = jm
      data1.Recordset!kolicina = text3(0)
      data1.Recordset!cijenarobe = text3(1)
      data1.Recordset!iznos = izn
   data1.Recordset.Update
   data1.Refresh
   brisiart
   vrati = False
   racunajtotal
Case 1
   If data1.Recordset.RecordCount = 0 Then Exit Sub
   brisiart
   If MsgBox("Jeste li sigurni da želite obrisati artikl: " & data1.Recordset!nazivrobe & "?", 36, "Provjera") = 7 Then Exit Sub
   data1.Recordset.Delete
   data1.Refresh
   vrati = False
   racunajtotal
Case 2
   If data1.Recordset.RecordCount = 0 Then Exit Sub
      na = ds(data1.Recordset!nazivrobe)
      jm = ds(data1.Recordset!jed_mjere)
      text3(0) = Format(db(data1.Recordset!kolicina), "standard")
      text3(1) = Format(db(data1.Recordset!cijenarobe), "standard")
      izn = Format(db(data1.Recordset!iznos), "standard")
      book = data1.Recordset.Bookmark
      vrati = True
Case 3
   brisiart
   vrati = False
End Select
na.SetFocus
Exit Sub
cmd_err1:
    greska
    Exit Sub
End Sub

Sub cmd1_Click (Index As Integer)
Dim i%
Select Case Index
Case 0
   If Not (Len(sk) > 0 And IsDate(maskededit1(0)) And IsDate(maskededit1(1)) And data1.Recordset.RecordCount > 0) Then MsgBox "Neispravno upisani podaci!", 16, "Greška": Exit Sub
   On Error GoTo rb
   BeginTrans
   Dim t As table
   Set t = b.OpenTable("ul_faktura")
   t.Index = "PrimaryKey"
   Dim mbr As Long, oldbr As Long
   If memode = 0 Then
      mbr = Maxbr_racun()
      oldbr = CLng(br)
      If oldbr <> mbr Then 'promjena broja raèuna ako je raèun pod brojem br veæ upisan
         br = mbr
         tg1.Active = False
         data1.Recordset.MoveFirst
         Do Until data1.Recordset.EOF
            data1.Recordset.Edit
            data1.Recordset!br_racun = mbr
            data1.Recordset.Update
         Loop
      End If
      t.AddNew
   Else
      t.Seek "=", memode
      mbr = memode
      t.Edit
   End If
      t!sifra_dob = sk
      If memode = 0 Then t!br_racun = mbr
      t!datum = maskededit1(0)
      t!valuta = maskededit1(1)
      t!racun = racun
      t!tip = cbotr
      t!uk_iznos = uk(0)
    For i = 0 To 3
        t(uk(i).DataField) = uk(i)
    Next i
    For i = 0 To 7
        t(txt(i).DataField) = txt(i)
    Next i
      If memode <> 0 Then
        'provjeri je li još uvijek plaæen
        If db(t!vsp) > db(t!placeno_izn) Then t!placeno = False
      End If

      t!nap = nap
   t.Update
'   If memode <> 0 Then 'prkopiraj iz pomoæne u glavnu bazu
   b.Execute "delete * from ul_roba where br_racun=" & mbr
   b.Execute "INSERT INTO ul_roba SELECT DISTINCTROW x_ul_roba.* FROM x_ul_roba where br_racun=" & br
'   End If
   CommitTrans
   If memode = 0 And oldbr <> mbr Then MsgBox "Broj raèuna je pri upisu promjenjen u " & mbr & " jer je raèun " & oldbr & "veæ upisan!", 64, "Promjena broja raèuna"
      mp 0
   Unload Me
Case 1
   'If memode = 0 Then
   If data1.Recordset.RecordCount = 0 Then Unload Me: Exit Sub
   If memode = 0 Then If MsgBox("Jeste li sigurni da želite obrisati raèun " & br & "?", 36, "Provjera") = 7 Then Exit Sub
   'On Error GoTo rb
   'BeginTrans
   '   If data1.Recordset.RecordCount > 0 Then b.Execute "delete * from ul_roba where br_racun= " & br
   'CommitTrans
   'Else
   BeginTrans
      If data1.Recordset.RecordCount > 0 Then b.Execute "delete * from x_ul_roba where br_racun= " & br
   CommitTrans
   'End If
   Unload Me
End Select
Exit Sub
rb:
   Rollback
   MsgBox Error, 16, "Greška"
   Exit Sub
End Sub

Sub Form_Load ()
BackColor = CLng(inidaj("Boje", "UR", Str(12632256)))
NamjestiTGBC Me
cbotiprac cbotr'puni combo tip racuna
Dim i%
top = 20
'left = 200
On Error Resume Next
data1.DatabaseName = PutBaza
memode = mode
Dim t As table
b.Execute "delete * from x_ul_roba where br_racun=" & memode
'puni combo jedinice mjere
   Set t = b.OpenTable("jed_mj")
   t.Index = "PrimaryKey"
   Do Until t.EOF
      jm.AddItem t!jed_mjere
      t.MoveNext
   Loop
Select Case memode
Case 0
   maskededit1(0) = df(Now)' Format(Now, "dd.mm.yyyy")
   br = Maxbr_racun()
   txt(4) = Format(25, "standard")
   'data1.RecordSource = "select * from ul_roba where br_racun=" & br
   pdvchg = False
Case Else
   Caption = "Ulazni raèuni - ispravak"
   Set t = b.OpenTable("ul_faktura")
   t.Index = "PrimaryKey"
   t.Seek "=", memode
   PrikaziKupca (t!sifra_dob)
   br = t!br_racun
   racun = t!racun
   maskededit1(0) = df(t!datum): maskededit1(1) = df(t!valuta): MaskedEdit1_LostFocus 0
   cbotiprac1 ds(t!tip), cbotr
   For i = 0 To 3
        uk(i) = Format(t(uk(i).DataField), "Standard")
   Next i
   For i = 0 To 7
        txt(i) = Format(t(txt(i).DataField), "Standard")
   Next i
   nap = t!nap
   b.Execute "INSERT INTO x_ul_roba SELECT DISTINCTROW ul_roba.* FROM ul_roba where br_racun=" & br
   pdvchg = True
End Select

   data1.RecordSource = "select * from x_ul_roba where br_racun=" & br
   data1.Refresh
   racunajtotal

CenterForm Me
top = 20
End Sub

Sub form_resize ()
If windowstate = 2 Then windowstate = 0
End Sub

Sub MaskedEdit1_KeyPress (Index As Integer, keyascii As Integer)
datkp maskededit1(Index), keyascii
If keyascii = 13 Then MaskedEdit1_LostFocus Index
End Sub

Sub MaskedEdit1_LostFocus (Index As Integer)
datlf maskededit1(Index)

If IsDate(maskededit1(0)) And IsDate(maskededit1(1)) Then
   text1(0) = CInt(CVDate(maskededit1(1)) - CVDate(maskededit1(0)))
Else
   text1(0) = ""
End If
End Sub

Function Maxbr_racun () As Long
   Dim t As table
   Set t = b.OpenTable("ul_faktura")
   t.Index = "PrimaryKey"
   If t.RecordCount > 0 Then t.MoveLast : Maxbr_racun = t!br_racun + 1 Else Maxbr_racun = 1
End Function

Sub na_KeyDown (KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 114
str1 = "select * from Qul_robapr where nazivrobe like '" & na & "*' order by nazivrobe"
Case 115
str1 = "select * from Qul_robapr where sifra_dob='" & sk & "' and nazivrobe like '" & na & "*' order by nazivrobe"
Case Else
    Exit Sub
End Select
Dim s As snapshot
   Set s = b.CreateSnapshot(str1)
   If s.RecordCount > 0 Then
      s.MoveNext
      If s.EOF Then
         s.MovePrevious
         str1 = s!nazivrobe
         str2 = s!sifra_dob
         s.Close
      Else
         s.Close
         mode = 3
         izbor.Show 1
      End If
      If str1 <> "" Then PrikaziRobu str1, str2, dbl1
   End If
End Sub

Sub por_Click (Index As Integer, Value As Integer)
racunajnaplatu 0
End Sub

Function porez (id%) As Double
Dim t As table
Set t = b.OpenTable("porezi")
t.Index = "PrimaryKey"
t.Seek "=", id
porez = t!porp
End Function

Sub PrikaziKupca (sku$)
On Error Resume Next
Dim t As table
Set t = b.OpenTable("partneri")
t.Index = "PrimaryKey"
t.Seek "=", sku
If Not t.NoMatch Then
   p.Caption = t!naziv_1
   p.Caption = p.Caption & Chr(13) & t!naziv_2
   p.Caption = p.Caption & Chr(13) & t!adresa & ", " & t!mjesto
   p.Caption = p.Caption & Chr(13) & "Porezni broj: " & t!por_br
   If IsNull(t!por_br) Or t!por_br = "" Then
        h "POREZNI BROJ nije upisan !!!"
        frm_mdi.Htimer.Interval = 5000
   End If
   cbotiprac1 ds(t!TipRacuna), cbotr
   sifk = sku
   sk = sifk
   maskededit1(0).SetFocus
Else
   sifk = ""
   p.Caption = ""
   sk = ""
End If
End Sub

Sub PrikaziRobu (n$, d$, c#)
On Error Resume Next
Dim t As table
Set t = b.OpenTable("ul_robapr")
t.Index = "PrimaryKey"
t.Seek "=", n, d, c
If Not t.NoMatch Then
   na = t!nazivrobe
   jm = t!jed_mjere
   text3(1) = Format(t!cijenarobe, "Standard")
   text3(0).SetFocus
Else
   na = ""
   jm = ""
   text3(1) = ""
   text3(0) = ""
   izn = ""
End If

End Sub

Sub racunajnaplatu (Index%)
Dim u() As Double, i%, t()  As Double
ReDim u(0 To 3)
ReDim t(0 To 7)

On Error Resume Next
For i = 0 To 3
    u(i) = cCDbl(uk(i))
Next i
For i = 0 To 7
    t(i) = cCDbl(txt(i))
Next i
On Error GoTo 0

If u(0) > 0 Then
    t(0) = t(1) * 100 / u(0)
    t(2) = t(3) * 100 / u(0)
End If
u(1) = u(0) - t(1) + t(3)

u(2) = zaokF(u(1) * t(4) / 100, 2)
If Index < 5 And Not pdvchg Then
    t(7) = u(2)
End If

u(3) = t(7) + u(1)

If t(6) = 0 Then
    t(5) = t(7)
    Else
        If t(6) = 0 Then
            t(6) = t(7)
        Else
            t(5) = t(7) - t(6)
        End If
End If

If t(7) <> t(5) + t(6) Then
    t(5) = t(7)
    t(6) = 0
End If

For i = 0 To 3
    uk(i) = Format(u(i), "standard")
Next i
For i = 0 To 7
    txt(i) = Format(t(i), "standard")
Next i


'Dim u#, r#, p#, ri#, pi#, zt#, zti#, ms#
'If IsNumeric(uk(0)) Then u = CDbl(uk(0)) Else u = 0
'If IsNumeric(text3(2)) Then ri = CDbl(text3(2)) Else If IsNumeric(text2(0)) Then ri = u * CDbl(text2(0)) / 100
'ms = u - ri
'If IsNumeric(text3(3)) Then pi = CDbl(text3(3)) Else If IsNumeric(text2(1)) Then pi = u * CDbl(text2(1)) / 100
'If IsNumeric(text3(4)) Then zti = CDbl(text3(4)) Else If IsNumeric(text2(2)) Then zti = u * CDbl(text2(2)) / 100
'If pi > ms Then pi = 0
''If zti > ms Then zti = 0
'If u > 0 Then text2(0) = Format(ri / u * 100, "standard") Else text2(0) = Format(0, "standard")
'If ms > 0 Then text2(1) = Format(pi / ms * 100, "standard") Else text2(1) = Format(0, "standard")
'If ms > 0 Then text2(2) = Format(zti / ms * 100, "standard") Else text2(2) = Format(0, "standard")
'text3(2) = Format(ri, "Standard")
'text3(3) = Format(pi, "Standard")
'text3(4) = Format(zti, "Standard")
'uk(3) = Format(ms + pi + zti, "standard")
End Sub

Sub racunajtotal ()
If data1.Recordset.RecordCount > 0 Then
   Dim s As snapshot, u#
   Set s = data1.Recordset.CreateSnapshot()
   Do
      'u = u + Format(s!kolicina * s!cijenarobe, "Standard")
      u = u + db(s!iznos)
      s.MoveNext
   Loop Until s.EOF
   uk(0) = Format(u, "standard")
   u = Format(u, "Standard")
Else
   uk(0) = Format(0, "Standard")
End If
racunajnaplatu 0
End Sub

Sub sk_KeyDown (KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode > 116 Or KeyCode < 114 Then Exit Sub
Dim s As snapshot


Select Case KeyCode
Case 114'po šifri
   str1 = "select * from partneri where dob and sifra_kupca like '" & sk & "*' order by sifra_kupca"
Case 115'po nazivu
   str1 = "select * from partneri where dob and naziv_1 like '" & sk & "*' order by naziv_1"
Case 116'po nazivu
   str1 = "select * from partneri where dob and naziv_2 like '" & sk & "*' order by naziv_2"
End Select

   Set s = b.CreateSnapshot(str1)
   If s.RecordCount > 0 Then
      s.MoveNext
      If s.EOF Then
         s.MovePrevious
         str1 = s!sifra_kupca
         s.Close
      Else
         s.Close
         mode = 1
         izbor.Show 1
      End If
      If str1 <> "" Then PrikaziKupca (str1)
   End If
End Sub

Sub sk_LostFocus ()
sk_KeyDown 114, 0
sk = sifk
End Sub

Sub Text1_KeyPress (Index As Integer, keyascii As Integer)
If keyascii = 13 Then Text1_LostFocus Index
End Sub

Sub Text1_LostFocus (Index As Integer)
intlf text1(0)
If IsNumeric(text1(0)) And IsDate(maskededit1(0)) Then
   maskededit1(1) = Format(CVDate(maskededit1(0)) + CInt(text1(0)), "dd.mm.yyyy")
Else
   maskededit1(1) = "__.__.____"
End If
End Sub

Sub Text2_KeyPress (Index As Integer, keyascii As Integer)
'keyascii = dblkp(keyascii)
'If keyascii = 13 Then Text2_LostFocus Index
End Sub

Sub Text2_LostFocus (Index As Integer)
'dbllf text2(Index)
''prijevod postotaka u iznose
'Dim u, u1, p As Double
'If IsNumeric(text2(Index)) Then p = CDbl(text2(Index)) / 100
'If IsNumeric(uk(0)) Then u = CDbl(uk(0))
'If Index <> 0 And IsNumeric(text3(2)) Then u = u - CDbl(text3(2))
'text3(Index + 2) = Format(u * p, "standard")
'RacunajNaplatu
End Sub

Sub Text3_KeyPress (Index As Integer, keyascii As Integer)
keyascii = dblkp(keyascii)
If keyascii = 13 Then Text3_LostFocus Index
End Sub

Sub Text3_LostFocus (Index As Integer)
dbllf text3(Index)
If Index < 2 Then
If IsNumeric(text3(0)) And IsNumeric(text3(1)) Then
   izn = Format(cCDbl(text3(0)) * cCDbl(text3(1)), "standard")
Else
   izn = ""
End If
Else
   racunajtotal
End If
End Sub

Sub txt_KeyPress (Index As Integer, keyascii As Integer)

keyascii = dblkp(keyascii)
If keyascii = 13 Then txt_LostFocus Index

End Sub

Sub txt_LostFocus (Index As Integer)
dbllf txt(Index)
If Index = 7 Then pdvchg = True
Dim d0#, d1#, r#, pdv#
'On Error Resume Next
d0 = cCDbl(uk(0))
d1 = cCDbl(uk(1))
pdv = cCDbl(txt(7))
'On Error GoTo 0

Select Case Index
Case 0, 2
    r = cCDbl(txt(Index)) / 100 * d0
    txt(Index + 1) = Format(r, "Standard")
Case 4
    r = cCDbl(txt(4)) / 100 * d1
    txt(7) = Format(r, "Standard")
Case 5
    txt(6) = Format(pdv - txt(5), "standard")
Case 6
    txt(5) = Format(pdv - txt(6), "standard")
End Select

racunajnaplatu Index
End Sub

