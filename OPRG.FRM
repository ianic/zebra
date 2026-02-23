VERSION 2.00
Begin Form oprg 
   BackColor       =   &H00C0C0C0&
   Caption         =   "O programu"
   ClientHeight    =   1665
   ClientLeft      =   5265
   ClientTop       =   3330
   ClientWidth     =   3645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   2175
   Left            =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3645
   Top             =   2880
   Width           =   3765
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   0
      Caption         =   "Programer: Igor Aniæ"
      Font3D          =   3  'Inset w/light shading
      Height          =   315
      Index           =   2
      Left            =   30
      TabIndex        =   4
      Top             =   750
      Width           =   3585
   End
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   0
      Caption         =   "Verzija: 2.00"
      Font3D          =   3  'Inset w/light shading
      Height          =   285
      Index           =   1
      Left            =   690
      TabIndex        =   3
      Top             =   450
      Width           =   2925
   End
   Begin SSPanel Panel3D2 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   0
      Caption         =   "Zebra - raèuni"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   405
      Index           =   0
      Left            =   690
      TabIndex        =   2
      Top             =   30
      Width           =   2925
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   0
      Font3D          =   3  'Inset w/light shading
      Height          =   585
      Left            =   30
      TabIndex        =   0
      Top             =   1080
      Width           =   3585
      Begin SSCommand Command3D1 
         AutoSize        =   2  'Adjust Button Size To Picture
         Caption         =   "Ok"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000080&
         Height          =   525
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   3525
      End
   End
   Begin Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   OPRG.FRX:0000
      Top             =   150
      Width           =   480
   End
End
Option Explicit

Sub Command3D1_Click ()
Unload Me
End Sub

Sub Form_Load ()
CenterForm Me
panel3d2(0).Caption = frm_mdi.Caption
End Sub

