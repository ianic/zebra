VERSION 2.00
Begin Form iz_rac 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Izlazni raèuni - novi"
   ClientHeight    =   6570
   ClientLeft      =   225
   ClientTop       =   2610
   ClientWidth     =   10860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   7080
   Icon            =   IZRAC.FRX:0000
   Left            =   165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   10860
   Top             =   2160
   Visible         =   0   'False
   Width           =   10980
   Begin SSPanel Panel3D9 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "oznaka"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   5610
      TabIndex        =   61
      Top             =   0
      Width           =   2955
      Begin ComboBox cboOZ 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   30
         Width           =   2295
      End
   End
   Begin SSPanel Panel3D8 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "naèin plaæanja"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   5610
      TabIndex        =   59
      Top             =   360
      Width           =   5235
      Begin ComboBox cboNPL 
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   30
         Width           =   4065
      End
   End
   Begin SSPanel pnlTR 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "tip"
      Font3D          =   3  'Inset w/light shading
      Height          =   375
      Left            =   10350
      TabIndex        =   58
      Top             =   0
      Width           =   495
   End
   Begin SSPanel pnlI 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "vsp"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   3
      Left            =   8640
      TabIndex        =   57
      Top             =   5400
      Width           =   1995
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
         Left            =   420
         TabIndex        =   52
         Top             =   30
         Width           =   1545
      End
   End
   Begin SSPanel pnlI 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "pdv"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   2
      Left            =   6660
      TabIndex        =   55
      Top             =   5400
      Width           =   1995
      Begin SSPanel uk 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         DataField       =   "pdv"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   2
         Left            =   420
         TabIndex        =   56
         Top             =   30
         Width           =   1545
      End
   End
   Begin SSPanel pnlI 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "rabat"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   1
      Left            =   4590
      TabIndex        =   53
      Top             =   5400
      Width           =   2085
      Begin SSPanel uk 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         DataField       =   "rab"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   1
         Left            =   510
         TabIndex        =   54
         Top             =   30
         Width           =   1545
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "csp"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   16
      Left            =   5550
      TabIndex        =   49
      Top             =   2010
      Width           =   1695
      Begin TextBox txt 
         BorderStyle     =   0  'None
         DataField       =   "csp"
         Height          =   225
         Index           =   5
         Left            =   360
         MaxLength       =   18
         TabIndex        =   12
         Top             =   30
         Width           =   1305
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
      Height          =   285
      Index           =   15
      Left            =   7230
      TabIndex        =   48
      Top             =   2010
      Width           =   1695
      Begin TextBox txt 
         BorderStyle     =   0  'None
         DataField       =   "vsp"
         Height          =   225
         Index           =   6
         Left            =   360
         MaxLength       =   18
         TabIndex        =   13
         Top             =   30
         Width           =   1305
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
      Height          =   285
      Index           =   12
      Left            =   2820
      TabIndex        =   46
      Top             =   2010
      Width           =   2745
      Begin SSPanel pnl 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   1410
         TabIndex        =   51
         Top             =   30
         Width           =   1275
      End
      Begin TextBox txt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "pdvp"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "23"
         Top             =   30
         Width           =   615
      End
      Begin SSPanel Panel3D4 
         Alignment       =   1  'Left Justify - MIDDLE
         AutoSize        =   1  'AutoSize Panel Width To Caption
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         Caption         =   "%  PDV"
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
         Left            =   810
         TabIndex        =   47
         Top             =   30
         Width           =   540
      End
   End
   Begin SSPanel Panel3D6 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   795
      Left            =   8970
      TabIndex        =   42
      Top             =   5790
      Width           =   855
      Begin MaskEdBox kopi 
         Height          =   285
         Left            =   540
         Mask            =   "#"
         MaxLength       =   1
         PromptChar      =   "_"
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   360
         Width           =   225
      End
      Begin SSPanel Panel3D7 
         Alignment       =   0  'Left Justify - TOP
         BevelOuter      =   0  'None
         BorderWidth     =   1
         Caption         =   "kopija"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   285
         Left            =   60
         TabIndex        =   44
         Top             =   390
         Width           =   465
      End
      Begin SSCheck ispis 
         Caption         =   "ispis"
         Font3D          =   3  'Inset w/light shading
         Height          =   255
         Left            =   60
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   90
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin SSPanel pnlS 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "S"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Symbol"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   1560
      TabIndex        =   41
      Top             =   5400
      Width           =   975
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
      TabIndex        =   40
      Top             =   5790
      Width           =   8925
      Begin TextBox nap 
         Height          =   735
         Left            =   840
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         Width           =   8085
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
      Height          =   285
      Index           =   11
      Left            =   30
      TabIndex        =   38
      Top             =   2010
      Width           =   2805
      Begin SSPanel pnl 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   1470
         TabIndex        =   50
         Top             =   30
         Width           =   1275
      End
      Begin SSPanel Panel3D4 
         Alignment       =   1  'Left Justify - MIDDLE
         AutoSize        =   1  'AutoSize Panel Width To Caption
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
         Left            =   810
         TabIndex        =   39
         Top             =   30
         Width           =   570
      End
      Begin TextBox txt 
         BorderStyle     =   0  'None
         DataField       =   "rabp"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   10
         Top             =   30
         Width           =   615
      End
   End
   Begin SSPanel pnlI 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "vr."
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   0
      Left            =   2610
      TabIndex        =   36
      Top             =   5400
      Width           =   1995
      Begin SSPanel uk 
         Alignment       =   4  'Right Justify - MIDDLE
         BevelOuter      =   0  'None
         BevelWidth      =   0
         BorderWidth     =   0
         DataField       =   "vbp"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   0
         Left            =   420
         TabIndex        =   37
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
      Left            =   9840
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5850
      Width           =   975
   End
   Begin SSCommand CMD1 
      Caption         =   "&Odustani"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   345
      Index           =   1
      Left            =   9840
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   6210
      Width           =   975
   End
   Begin Data Data1 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   "C:\VB\DATA\R97\ZEBRA_97.MDB"
      Exclusive       =   0   'False
      Height          =   285
      Left            =   9510
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   "prodroba"
      Top             =   2220
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
      Left            =   6960
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2340
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
      Left            =   4650
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2340
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
      Left            =   2340
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2340
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
      Left            =   30
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2340
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
      TabIndex        =   28
      Top             =   1650
      Width           =   1755
      Begin ComboBox jm 
         Height          =   315
         Left            =   600
         TabIndex        =   6
         Tag             =   "jed_mjere"
         Top             =   30
         Width           =   1125
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "vr."
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   8
      Left            =   7230
      TabIndex        =   27
      Top             =   1740
      Width           =   1695
      Begin TextBox txt 
         BorderStyle     =   0  'None
         DataField       =   "vbp"
         Height          =   225
         Index           =   2
         Left            =   360
         MaxLength       =   18
         TabIndex        =   9
         Top             =   30
         Width           =   1305
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
      Height          =   285
      Index           =   7
      Left            =   5370
      TabIndex        =   26
      Top             =   1740
      Width           =   1875
      Begin TextBox txt 
         BorderStyle     =   0  'None
         DataField       =   "cbp"
         Height          =   225
         Index           =   1
         Left            =   540
         MaxLength       =   18
         TabIndex        =   8
         Top             =   30
         Width           =   1305
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
      Height          =   285
      Index           =   6
      Left            =   1770
      TabIndex        =   25
      Top             =   1740
      Width           =   1995
      Begin TextBox txt 
         BorderStyle     =   0  'None
         DataField       =   "kolicina"
         Height          =   225
         Index           =   0
         Left            =   660
         MaxLength       =   18
         TabIndex        =   7
         Top             =   0
         Width           =   1305
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
      Height          =   375
      Index           =   5
      Left            =   30
      TabIndex        =   24
      Top             =   1290
      Width           =   8175
      Begin SSPanel Panel3D2 
         Alignment       =   0  'Left Justify - TOP
         BevelOuter      =   0  'None
         BorderWidth     =   1
         Caption         =   "F4 - naziv artikla"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   6870
         TabIndex        =   29
         Top             =   90
         Width           =   1245
      End
      Begin TextBox na 
         DataField       =   "nazivrobe"
         Height          =   300
         Left            =   960
         MaxLength       =   50
         TabIndex        =   5
         Top             =   30
         Width           =   5865
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "rok plaæanja                 dana"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   645
      Index           =   2
      Left            =   8520
      TabIndex        =   22
      Top             =   720
      Width           =   2325
      Begin TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1020
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   30
         Width           =   585
      End
      Begin MaskEdBox MaskedEdit1 
         Height          =   285
         Index           =   1
         Left            =   1020
         Mask            =   "##/##/####"
         MaxLength       =   10
         PromptChar      =   "_"
         TabIndex        =   4
         Top             =   330
         Width           =   1245
      End
      Begin SSPanel Panel3D3 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelOuter      =   0  'None
         BorderWidth     =   1
         Caption         =   "valuta"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   23
         Top             =   330
         Width           =   495
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
      Left            =   6570
      TabIndex        =   21
      Top             =   720
      Width           =   1965
      Begin MaskEdBox MaskedEdit1 
         Height          =   285
         Index           =   0
         Left            =   690
         Mask            =   "##/##/####"
         MaxLength       =   10
         PromptChar      =   "_"
         TabIndex        =   2
         Top             =   30
         Width           =   1245
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
      Height          =   375
      Index           =   0
      Left            =   8550
      TabIndex        =   19
      Top             =   0
      Width           =   1815
      Begin SSPanel br 
         BevelOuter      =   0  'None
         BorderWidth     =   1
         Caption         =   "12345"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   690
         TabIndex        =   20
         Top             =   60
         Width           =   1065
      End
   End
   Begin SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "kupac"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   345
      Left            =   60
      TabIndex        =   17
      Top             =   0
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
         Left            =   1950
         TabIndex        =   18
         Top             =   60
         Width           =   2835
      End
      Begin TextBox sk 
         Height          =   285
         Left            =   570
         MaxLength       =   10
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
      Height          =   885
      Left            =   60
      TabIndex        =   16
      Top             =   360
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
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeadingHeight   =   1
      Height          =   2745
      HorzLines       =   2  '3D
      Layout          =   IZRAC.FRX:0302
      LayoutIndex     =   1
      Left            =   0
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2640
      UseBookmarks    =   -1  'True
      Width           =   10845
      WrapCellPointer =   0   'False
   End
End
Option Explicit
Dim sifk$, memode%, unl%
Dim vrati%, book

Sub brisiart ()
   na = ""
   jm = ""
   txt(0) = ""
   txt(1) = ""
   txt(2) = ""
   'txt(3) = ""
   txt(5) = ""
   txt(6) = ""
   If cbonpl.ListIndex >= 2 Then txt(4) = "0" Else txt(4) = inidaj("Baze", "PdvPOstotak", "22")
   pnl(1) = ""
   pnl(2) = ""

  ' text3(0) = ""
  ' text3(1) = ""
   'izn = ""
End Sub

Sub cboNPL_Click ()
brisiart
End Sub

Sub CMD_Click (Index As Integer)
On Error GoTo cmd_error
Dim i%
Select Case Index
Case 0
   txt_LostFocus 1
   If Not (Len(na) > 0 And IsNumeric(txt(6)) And Len(jm.Text) > 0) Then MsgBox "Neispravno upisani podaci!", 16, "Greška": Exit Sub
   If vrati Then data1.Recordset.Bookmark = book: data1.Recordset.Edit  Else data1.Recordset.AddNew
      data1.Recordset!br_racun = br
      data1.Recordset!nazivrobe = na
      data1.Recordset!jed_mjere = jm
      data1.Recordset!rabi = pnl(1)
      data1.Recordset!pdvi = pnl(2)
      For i = 0 To 6
        data1.Recordset(txt(i).DataField) = txt(i)
      Next i
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
      br = data1.Recordset!br_racun
      na = data1.Recordset!nazivrobe
      jm = data1.Recordset!jed_mjere

      pnl(1) = data1.Recordset!rabi
      pnl(2) = data1.Recordset!pdvi
      For i = 0 To 6
        txt(i) = data1.Recordset(txt(i).DataField)
      Next i
      vrati = True
      book = data1.Recordset.Bookmark
      txt_LostFocus 1
Case 3
    brisiart
    vrati = False
End Select
na.SetFocus
Exit Sub
cmd_error:
    greska
    Exit Sub
End Sub

Sub cmd1_Click (Index As Integer)
Dim i%
mp 11
Select Case Index
Case 0
   If Not (Len(sk) > 0 And IsDate(maskededit1(0)) And IsDate(maskededit1(1)) And data1.Recordset.RecordCount > 0) Then MsgBox "Neispravno upisani podaci!", 16, "Greška": mp 0: Exit Sub
   On Error GoTo rb
   BeginTrans
   Dim t As table
   Set t = b.OpenTable("faktura")
   t.Index = "PrimaryKey"
   Dim mbr As Long, oldbr As Long
   If memode = 0 Then
      mbr = Maxbr_racun()
      oldbr = CDbl(br)
      If oldbr <> mbr Then 'promjena broja raèuna ako je raèun pod brojem br veæ upisan
         br = mbr
         tg1.Active = False
         data1.Recordset.MoveFirst
         Do Until data1.Recordset.EOF
            data1.Recordset.Edit
            data1.Recordset!br_racun = mbr
            data1.Recordset.Update
            data1.Recordset.MoveNext
         Loop
      End If
      t.AddNew
   Else
      t.Seek "=", memode
      mbr = memode
      t.Edit
   End If
      t!sifra_kupca = sk
      If memode = 0 Then t!br_racun = mbr
      t!datum = maskededit1(0)
      t!valuta = maskededit1(1)
      t!tip = pnlTR
      For i = 0 To 3
        t(uk(i).DataField) = uk(i)
      Next i
      t!npl = IIf(cbonpl.ListIndex > 0, cbonpl.ListIndex, 0)
      t!oznaka = Left(cbooz, 1)
      t!nap = nap
      If memode <> 0 Then
        'provjeri je li još uvijek plaæen
        If db(t!vsp) > db(t!placeno_izn) Then t!placeno = False
      End If
   t.Update
   'If memode <> 0 Then 'prkopiraj iz pomoæne u glavnu bazu
   b.Execute "delete * from prodroba where br_racun=" & mbr
   b.Execute "INSERT INTO prodroba SELECT DISTINCTROW x_prodroba.* FROM x_prodroba where br_racun=" & br
   'End If
   CommitTrans
   If memode = 0 And oldbr <> mbr Then MsgBox "Broj raèuna je pri upisu promjenjen u " & mbr & " jer je raèun " & oldbr & "veæ upisan!", 64, "Promjena broja raèuna"
   If ispis Then
   '   mp 11 ispis raèuna
      Dim k%
      If IsNumeric(kopi) Then
        k = CInt(kopi)
      Else
        k = 1
      End If
      If CInt(inidaj("Ispis", "KopijeAuto", "0")) Then
            ispisracuna mbr, 0, 3
      Else
         For i = 1 To k
            ispisracuna mbr, 0, 1
         Next i
      End If
   '   mp 0
   End If
   Unload Me
Case 1
   'If memode = 0 Then
   If data1.Recordset.RecordCount = 0 Then Unload Me: mp 0: Exit Sub
   If memode = 0 Then
    If MsgBox("Jeste li sigurni da želite obrisati raèun " & br & "?", 36, "Provjera") = 7 Then Exit Sub
   End If
   'On Error GoTo rb
   'BeginTrans
   '   If data1.Recordset.RecordCount > 0 Then b.Execute "delete * from prodroba where br_racun= " & br
   'CommitTrans
   'Else
   BeginTrans
      If data1.Recordset.RecordCount > 0 Then b.Execute "delete * from x_prodroba where br_racun= " & br
   CommitTrans
   'End If
   Unload Me
End Select
mp 0
Exit Sub
rb:
    mp 0
    On Error Resume Next
    Rollback
    greska
    Exit Sub
End Sub

Sub Form_Load ()
pnlTR = inidaj("IR", "TIP", "R-2")
BackColor = CLng(inidaj("Boje", "IR", Str(8421376)))
NamjestiTGBC Me

Dim i%
top = 100
left = 100
On Error Resume Next
data1.DatabaseName = PutBaza
memode = mode
Dim t As table
cboNplac cbonpl 'combo naèin plaæanja
cboOznaka cbooz
   'napuni combo jedinice mjere
   Set t = b.OpenTable("jed_mj")
   t.Index = "PrimaryKey"
   Do Until t.EOF
      jm.AddItem t!jed_mjere
      t.MoveNext
   Loop

b.Execute "delete * from x_prodroba where br_racun=" & memode

Select Case memode
Case 0'novi raèun
   ispis = True: kopi = 3
   maskededit1(0) = df(Now): maskededit1(1) = df(Now + 15): text1(0) = 15
   br = Maxbr_racun()
   'data1.RecordSource = "select * from prodroba where br_racun=" & br
   top = 100
Case Else'postojeæi raèun
   ispis = False: kopi = 1
   Caption = "Izlazni raèuni - ispravak"
   Set t = b.OpenTable("faktura")
   t.Index = "PrimaryKey"
   t.Seek "=", memode
   PrikaziKupca (t!sifra_kupca)
   br = t!br_racun
   maskededit1(0) = df(ds(t!datum)): maskededit1(1) = df(ds(t!valuta)): MaskedEdit1_LostFocus 0
   pnlTR = ds(t!tip)
   cbonpl.ListIndex = db(t!npl)
   cboOZNAKA1 ds(t!oznaka), cbooz
   nap = t!nap
   b.Execute "INSERT INTO x_prodroba SELECT DISTINCTROW prodroba.* FROM prodroba where br_racun=" & br
End Select

   data1.RecordSource = "select * from x_prodroba where br_racun=" & br
   data1.Refresh
   racunajtotal

'CenterForm Me
End Sub

Sub form_resize ()
If windowstate = 2 Then windowstate = 0
End Sub

Sub MaskedEdit1_KeyPress (Index As Integer, KeyAscii As Integer)
datkp maskededit1(Index), KeyAscii
If KeyAscii = 13 Then MaskedEdit1_LostFocus Index
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
   Set t = b.OpenTable("faktura")
   t.Index = "PrimaryKey"
   If t.RecordCount > 0 Then t.MoveLast : Maxbr_racun = t!br_racun + 1 Else Maxbr_racun = 1
End Function

Sub na_GotFocus ()
cbooz.TabStop = False
maskededit1(0).TabStop = False
maskededit1(1).TabStop = False
text1(0).TabStop = False
End Sub

Sub na_KeyDown (KeyCode As Integer, Shift As Integer)
If KeyCode <> 115 Then Exit Sub
Dim s As snapshot
str1 = "select * from roba where nazivrobe like '" & na & "*' and grupa like '" & Left(cbooz, 1) & "*'" '"*' order by nazivrobe"
   Set s = b.CreateSnapshot(str1)
   If s.RecordCount > 0 Then
      s.MoveNext
      If s.EOF Then
         s.MovePrevious
         str1 = s!nazivrobe
         str2 = s!grupa
         s.Close
      Else
         s.Close
         mode = 2
         izbor.Show 1
      End If
      If str1 <> "" Then PrikaziRobu str1, str2
   End If
End Sub

Sub pnlS_Click ()

'If pnls.ForeColor = &H80& Then
'    pnls.ForeColor = &H800000
'Else
'    pnls.ForeColor = &H80&
'End If
racunajtotal
End Sub

Sub por_Click (Index As Integer, Value As Integer)
racunajnaplatu
End Sub

Function porez (id%) As Double
Dim t As table
Set t = b.OpenTable("porezi")
t.Index = "PrimaryKey"
t.Seek "=", id
porez = t!porp
End Function

Sub PrikaziKupca (sku$)
'On Error Resume Next
Dim t As snapshot'table
Set t = b.CreateSnapshot("select * from partneri where sifra_kupca='" & sku & "'")
't.Index = "PrimaryKey"
't.Seek "=", sku
'If Not t.NoMatch Then
If t.RecordCount <> 0 Then
   p.Caption = t!naziv_1
   p.Caption = p.Caption & Chr(13) & t!naziv_2
   p.Caption = p.Caption & Chr(13) & t!adresa & ", " & t!mjesto
   p.Caption = p.Caption & Chr(13) & "Porezni broj: " & t!por_br
   If IsNull(t!por_br) Or t!por_br = "" Then
        h "POREZNI BROJ nije upisan !!!"
        frm_mdi.Htimer.Interval = 5000
   End If
   sifk = sku
   sk = sifk
   maskededit1(0).SetFocus
   sk.TabStop = False
Else
   sifk = ""
   p.Caption = ""
   sk = ""
End If
End Sub

Sub PrikaziRobu (naz$, gru$)
'On Error Resume Next
Dim t As table
Set t = b.OpenTable("roba")
t.Index = "PrimaryKey"
t.Seek "=", gru, naz
If Not t.NoMatch Then
   na = t!nazivrobe
   jm = t!jed_mjere
   txt(1) = Format(db(t!cijenarobe), "standard")
   txt(4) = Format(db(t!pdvp), "standard")
   txt(5) = Format(db(t!csp), "standard")
   txt(2) = ""
   txt(3) = ""
   txt(6) = ""
   txt(0) = 1

   txt(0).SetFocus
Else
    brisiart
   'na = ""
   'jm = ""
   'text3(1) = ""
   'text3(0) = ""
   'izn = ""
End If

End Sub

Sub racunajnaplatu ()
'Exit Sub
'Dim u#, r#, p#, ri#, pi#
'If IsNumeric(uk(0)) Then u = CDbl(uk(0)) Else u = 0
'If IsNumeric(text2(0)) Then r = CDbl(text2(0)) Else r = 0
''If IsNumeric(text2(1)) Then p = CDbl(text2(1)) Else p = 0
'If por(0) Then p = 15
'If por(1) Then p = 10
'Dim i%
'For i = 0 To 5
'   If por(i) Then p = porez(i)
'Next i
'ri = Format(u * r / 100, "standard")
'uk(1) = Format(u * r / 100, "standard")
'ri = u - ri
'uk(2) = Format(ri * p / 100, "standard")
'pi = Format(ri * p / 100, "standard")
'uk(3) = Format(ri + pi, "standard")
End Sub

Sub racunajtotal ()
Dim u() As Double, i%, pdv22%
ReDim u(0 To 3)
pdv22 = True
If data1.Recordset.RecordCount > 0 Then
   Dim s As snapshot
   Set s = data1.Recordset.CreateSnapshot()
   Do
      'u(0) = u(0) + zaokf(db(s!vbp), 2)
      u(1) = u(1) + zaokf(db(s!rabi), 2)
      If db(s!pdvp) <> 22 Then pdv22 = False
      u(2) = u(2) + zaokf(db(s!pdvi), 2)
      u(3) = u(3) + zaokf(db(s!vsp), 2)
      s.MoveNext
   Loop Until s.EOF
End If

If pdv22 Then 'pdv je na svim stavkama 22%
   u(0) = zaokf((u(3) + u(1)) * .22, 2)
   pnls.ForeColor = &H80&
Else
    pnls.ForeColor = &H800000
End If

u(0) = zaokf(u(3) + u(1) - u(2), 2)

For i = 0 To 3
    uk(i) = Format(u(i), "Standard")
    uk(i).ForeColor = pnls.ForeColor
    pnlI(i).ForeColor = pnls.ForeColor
Next i

              

'If data1.Recordset.RecordCount > 0 Then
'   Dim s As snapshot, u#
'   Set s = data1.Recordset.CreateSnapshot()
'   Do
'      u = u + Format(s!kolicina * s!cijenarobe, "Standard")
'      s.MoveNext
'   Loop Until s.EOF
'   uk(0) = Format(u, "standard")
'   u = Format(u, "Standard")
'Else
'   uk(0) = Format(0, "Standard")
'End If
'RacunajNaplatu
End Sub

Sub sk_KeyDown (KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode > 116 Or KeyCode < 114 Then Exit Sub
Dim s As snapshot

Select Case KeyCode
Case 114'po šifri
   str1 = "select * from partneri where kup and sifra_kupca like '" & sk & "*' order by sifra_kupca"
Case 115'po nazivu
   str1 = "select * from partneri where kup and naziv_1 like '" & sk & "*' order by naziv_1"
Case 116'po nazivu
   str1 = "select * from partneri where kup and naziv_2 like '" & sk & "*' order by naziv_2"
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

Sub Text1_KeyPress (Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Text1_LostFocus Index
End Sub

Sub Text1_LostFocus (Index As Integer)
intlf text1(0)
If IsNumeric(text1(0)) And IsDate(maskededit1(0)) Then
    maskededit1(1) = df(CVDate(maskededit1(0)) + CDbl(text1(0)))
Else
   maskededit1(1) = "__.__.____"
End If
End Sub

Sub Text2_KeyPress (Index As Integer, KeyAscii As Integer)
KeyAscii = dblkp(KeyAscii)
If KeyAscii = 13 Then Text2_LostFocus Index
End Sub

Sub Text2_LostFocus (Index As Integer)
'dbllf text2(Index)
'RacunajNaplatu
End Sub

Sub Text3_KeyPress (Index As Integer, KeyAscii As Integer)
'keyascii = dblkp(keyascii)
'If keyascii = 13 Then Text3_LostFocus Index
End Sub

Sub Text3_LostFocus (Index As Integer)
'dbllf text3(Index)
'If IsNumeric(text3(0)) And IsNumeric(text3(1)) Then
'   izn = Format(CDbl(text3(0)) * CDbl(text3(1)), "standard")
'Else
'   izn = ""
'End If
End Sub

Sub txt_GotFocus (Index As Integer)
sel
End Sub

Sub txt_KeyPress (Index As Integer, KeyAscii As Integer)
KeyAscii = dblkp(KeyAscii)
If KeyAscii = 13 Then txt_LostFocus Index
End Sub

Sub txt_LostFocus (Index As Integer)
dbllf txt(Index)
Dim d() As Double, vbp#, rabi#, pori#, i%
ReDim d(0 To 6)
For i = 0 To 6
    d(i) = DajBrTxt((txt(i)))
Next i
'*0-kolièina
'*1-cbp
'2-vr.          =0*1
'*3-%rabat
'*4-%PDV
'5-csp          =vsp/kol
'6-vsp          =vr.+rabi+pdvi
'priprema za proraèun
If d(0) = 0 Then Exit Sub

Select Case Index
 Case 0, 1   ' u originalu
     d(2) = d(1) * d(0)    ' u originalu
     d(6) = d(6) * (1 - d(3) / 100) * (1 + d(4) / 100)   ' u originalu
     d(5) = d(6) / d(0)   ' u originalu

Case 2
    d(1) = d(2) / d(0)

Case 5, 6
    If Index = 5 Then d(6) = d(5) * d(0)

    rabi = vbp * d(3) / 100       ' u originalu
    pori = (vbp - rabi) * d(4) / 100     ' u originalu

    d(2) = d(6) / (1 - d(3) / 100) / (1 + d(4) / 100)  'd(6) + rabi - pori
    d(1) = d(2) / d(0)
End Select

'proraèun
 Select Case Index   ' u originalu
 Case 0, 1, 2     ' u originalu
    zaok d(1), 2
    d(2) = d(1) * d(0)
    vbp = d(2)
    rabi = vbp * d(3) / 100
    pori = (vbp - rabi) * d(4) / 100

    zaok d(2), 2    ' u originalu
    zaok rabi, 4
    zaok pori, 4

    d(6) = d(2) - rabi + pori
    d(5) = d(6) / d(0)

 Case 3, 4, 5, 6           ' u originalu
     d(5) = d(6) / d(0)     ' u originalu
     vbp = d(6) / (1 + d(4) / 100) / (1 - d(3) / 100)       ' u originalu
     vbp = d(6) / (1 - d(3) / 100 + d(4) / 100 - d(3) * d(4) / 10000)    ' u originalu
     rabi = vbp * d(3) / 100               ' u originalu
     pori = (vbp - rabi) * d(4) / 100      ' u originalu

     zaok d(6), 2      ' u originalu
     zaok rabi, 2     ' u originalu
     zaok pori, 2      ' u originalu

     d(2) = d(6) + rabi - pori     ' u originalu
     d(1) = d(2) / d(0)            ' u originalu
 End Select      ' u originalu

For i = 1 To 6
    txt(i) = Format(d(i), "#,##0.00##")
Next i

pnl(1) = Format(rabi, "#,##0.00##")
pnl(2) = Format(pori, "#,##0.00##")

End Sub

