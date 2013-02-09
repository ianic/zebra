VERSION 2.00
Begin Form ispis1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Ispis"
   ClientHeight    =   975
   ClientLeft      =   2235
   ClientTop       =   1695
   ClientWidth     =   4755
   Height          =   1485
   Left            =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4755
   Top             =   1245
   Width           =   4875
   Begin SSCommand Command3D1 
      AutoSize        =   2  'Adjust Button Size To Picture
      Caption         =   "Ispis"
      Font3D          =   3  'Inset w/light shading
      Height          =   960
      Left            =   3480
      Picture         =   ISPIS.FRX:0000
      TabIndex        =   2
      Top             =   0
      Width           =   1230
   End
   Begin SSOption o 
      Caption         =   "ispis raèuna broj: "
      Font3D          =   3  'Inset w/light shading
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   510
      Width           =   3045
   End
   Begin SSOption o 
      Caption         =   "ispis svih prikazanih raèuna"
      Font3D          =   3  'Inset w/light shading
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Value           =   -1  'True
      Width           =   3015
   End
End
Option Explicit

Sub Command3D1_Click ()
If o(0) Then MODE = 0
If o(1) Then MODE = 1
Unload Me
End Sub

Sub Form_Load ()
CenterForm Me
o(0).Caption = o(0).Caption & str1
MODE = 3
End Sub

