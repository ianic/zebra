VERSION 2.00
Begin Form tgpr 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Podešavanje/odabir ispisa"
   ClientHeight    =   7020
   ClientLeft      =   645
   ClientTop       =   1440
   ClientWidth     =   6255
   ForeColor       =   &H00808080&
   Height          =   7530
   KeyPreview      =   -1  'True
   Left            =   585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   6255
   Top             =   990
   Width           =   6375
   Begin SSPanel pnl3 
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Height          =   315
      Left            =   30
      TabIndex        =   16
      Top             =   5910
      Width           =   6195
      Begin SpinButton sp 
         Height          =   240
         Left            =   90
         SpinOrientation =   1  'Horizontal
         Top             =   60
         Width           =   495
      End
   End
   Begin TextBox mask 
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   13
      Tag             =   "mask"
      Text            =   "99"
      Top             =   900
      Width           =   315
   End
   Begin SSCommand cmd1 
      Caption         =   "&Ispis"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   4140
      Picture         =   TGPR.FRX:0000
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1050
   End
   Begin SSCommand cmd2 
      Caption         =   "&Odustani"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   795
      Left            =   5190
      Picture         =   TGPR.FRX:021C
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1050
   End
   Begin SSCheck chksum 
      Caption         =   "S"
      Font3D          =   3  'Inset w/light shading
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Symbol"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   1
      Left            =   2190
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   870
      Width           =   465
   End
   Begin SSPanel Panel3D2 
      Alignment       =   0  'Left Justify - TOP
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Caption         =   "Naziv konfiguracije"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   6225
      Begin SSCheck di 
         Caption         =   "datum ispisa"
         Font3D          =   3  'Inset w/light shading
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   570
         Value           =   -1  'True
         Width           =   1185
      End
      Begin TextBox Text2 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   14
         Top             =   600
         Width           =   4485
      End
      Begin SSPanel Panel3D1 
         Alignment       =   1  'Left Justify - MIDDLE
         BevelOuter      =   0  'None
         BorderWidth     =   2
         Caption         =   "zaglavlje"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   900
         TabIndex        =   10
         Top             =   330
         Width           =   795
      End
      Begin TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   4485
      End
      Begin SSCheck snimi 
         Caption         =   "snimi"
         Font3D          =   3  'Inset w/light shading
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5400
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   30
         Value           =   -1  'True
         Width           =   795
      End
      Begin ComboBox Combo1 
         Height          =   315
         Left            =   1710
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   30
         Width           =   3615
      End
   End
   Begin SSCheck check 
      Caption         =   "Check3D1"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   870
      Value           =   -1  'True
      Width           =   3015
   End
   Begin SSPanel panf2 
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   6630
      Width           =   585
   End
   Begin SSPanel panf1 
      BevelOuter      =   1  'Inset
      BorderWidth     =   2
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   6330
      Width           =   1905
   End
   Begin CommonDialog CMDialog1 
      Left            =   2400
      Top             =   6510
   End
   Begin SSCommand cmd4 
      Caption         =   "&Font"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   0
      Picture         =   TGPR.FRX:03C4
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1050
   End
   Begin SSCommand cmd3 
      Caption         =   "&Pisaè"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00808000&
      Height          =   795
      Left            =   3030
      Picture         =   TGPR.FRX:07FE
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1050
   End
End
Option Explicit

Sub check_Click (Index As Integer, Value As Integer)
ProracunajSirinu
End Sub

Sub cmd1_Click ()
If combo1 = "" And snimi Then
    MsgBox "Upišite naziv konfiguracije ukoliko želite snimiti konfiguraciju!!!", 32, ""
    Exit Sub
End If
Tag = ""
Hide
End Sub

Sub cmd2_click ()
Tag = "1"
Hide
End Sub

Sub cmd3_Click ()
frmPrnOdabir.Show 1
printer.EndDoc
'cmdialog1.Action = 5
DoEvents
panf1.Caption = printer.FontName
panf2.Caption = printer.FontSize

End Sub

Sub cmd4_Click ()
On Error Resume Next
cmdialog1.Flags = &H2&
cmdialog1.Action = 4
'pokušaj dodati "CE" na kraj
printer.FontName = cmdialog1.FontName & " CE"
If Err <> 0 Then
    printer.FontName = cmdialog1.FontName
    Err = 0
End If
printer.FontSize = cmdialog1.FontSize

panf1.Caption = printer.FontName
panf2.Caption = printer.FontSize
ProracunajSirinu
End Sub

Sub Combo1_Click ()
If IsNumeric(combo1.Tag) Then punitgprizbaze CInt(combo1.Tag), combo1.Text
End Sub

Sub Form_Activate ()
ProracunajSirinu
End Sub

Sub Form_KeyPress (keyascii As Integer)
If keyascii = 27 Then cmd2_click
End Sub

Sub Form_Load ()
'centerform Me
'mp 0
Tag = "1"
End Sub

Sub mask_GotFocus (Index As Integer)
sel
End Sub

Sub mask_LostFocus (Index As Integer)
If Not IsNumeric(mask(Index)) Then
    mask(Index) = 0
    check(Index) = False
Else
    If mask(Index) = "0" Then check(Index) = False
End If
ProracunajSirinu
End Sub

Sub ProracunajSirinu ()
Dim l&, i%, lmax&
    l = 0
    For i = 0 To controls.Count - 1
        If controls(i).Tag = "mask" Then
        If check(controls(i).Index) Then
        If draftispis Then
            l = l + CInt(controls(i)) + 2
        Else
            l = l + DajSirinuKolone((controls(i)), (check(controls(i).Index).Caption))
        End If
        End If
        End If
    Next i
If draftispis Then
    pnl3.Caption = l & " znakova od 160 moguæih"
    lmax = 160
Else
    pnl3.Caption = l & " toèaka od " & printer.ScaleWidth & " moguæih"
    lmax = printer.ScaleWidth
End If

If l > lmax Then
    pnl3.ForeColor = &H80& 'crveno
Else
    pnl3.ForeColor = 0 'crno
End If
End Sub

Sub sizemask (Index#)
Dim i%, j%
For i = 0 To controls.Count - 1
Debug.Print controls(i).Tag
    If Trim(controls(i).Tag) = "mask" Then
    
    If IsNumeric(controls(i)) Then
        Debug.Print controls(i)
        If Not (CInt(controls(i)) = 0 And Index < 1) Then
        j = Int(CInt(controls(i)) * Index)
        j = IIf(j = CInt(controls(i)), j + 1, j)
        If j > 99 Then j = 99
        controls(i) = j
        If j = 0 Then check(controls(i).Index) = False
        Debug.Print controls(i)
        End If
    End If
    End If
Next i
ProracunajSirinu
End Sub

Sub sp_SpinDown ()
sizemask .9
End Sub

Sub sp_SpinUp ()
sizemask 1.1
End Sub

