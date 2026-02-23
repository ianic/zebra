VERSION 2.00
Begin Form frmKalkulator 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Zbrajalica"
   ClientHeight    =   6540
   ClientLeft      =   1155
   ClientTop       =   960
   ClientWidth     =   6960
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "Courier"
   FontSize        =   12
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   ForeColor       =   &H00C0C000&
   Height          =   7050
   Icon            =   CALCULAT.FRX:0000
   Left            =   1095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   6960
   Top             =   510
   Width           =   7080
   Begin CommandButton cmdIspis 
      Caption         =   "Ispis"
      Height          =   345
      Left            =   5940
      TabIndex        =   6
      Top             =   450
      Visible         =   0   'False
      Width           =   915
   End
   Begin TextBox T2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   5625
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   840
      Width           =   3435
   End
   Begin TextBox tDec 
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "2"
      Top             =   360
      Width           =   255
   End
   Begin CommandButton cBrisi 
      Cancel          =   -1  'True
      Caption         =   "Briši sve"
      Height          =   375
      Left            =   5940
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin SSPanel lRez 
      BevelWidth      =   2
      BorderWidth     =   4
      Caption         =   " "
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   765
      Left            =   570
      TabIndex        =   2
      Top             =   30
      Width           =   5295
   End
   Begin TextBox t 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "System"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   5625
      Left            =   3510
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   3315
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dec."
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   420
   End
End
Option Explicit
  Dim dec% 'broj decimala

Sub Cbrisi_Click ()
   t = ""
   t.SetFocus
End Sub

Sub cmdIspis_Click ()
   Dim s$, i%, c%, p%, red() As Double
   Dim suma#, temp$
   ReDim red(200)

   c = 0
   s = Trim(t)
   While Len(s) > 0
      c = c + 1
      p = InStr(s, Chr(13) & Chr$(10)): If p = 0 Then p = 200
      red(c) = Val(Left(s, p))
      s = Mid(s, p + 2)
   Wend
''''   OtvoriPrn 5, 0, 72
      For i = 1 To c
         If red(i) < 0 Then bold 2 Else bold 0
         Iob Format(Abs(red(i)), "#,##0." & String(dec, "0")), 30, 1:
         Iob IIf(red(i) > 0, "+", "-"), 1, 1
         NLine
         suma = suma + red(i)
      Next
      Iob "-", 31, 5: NLine
      bold 3
      Iob Format(suma, "#,##0." & String(dec, "0")), 30, 1
      Iob "=", 1, 1
      NLine
   Close
End Sub

Sub Form_Click ()
   Zbroji
End Sub

Sub Form_Load ()
   Me.Show
   tDec_Change
   t.SetFocus
End Sub

Sub t_Change ()
Zbroji
End Sub

Sub t_KeyPress (keyascii As Integer)
   'Debug.Print KEYASCII
   If keyascii = 43 Then keyascii = 13 ' + -> enter
   If keyascii = 44 Then keyascii = 46 ' , -> .
   If keyascii <> 13 And keyascii <> 8 Then
      If keyascii < 43 Or keyascii > 57 Then keyascii = 0
      If keyascii = 44 Or keyascii = 47 Then keyascii = 0
   End If
End Sub

Sub T2_KeyPress (keyascii As Integer)
   keyascii = 0
End Sub

Sub tDec_Change ()
   
   tDec = IIf(Val(tDec) > 0, Val(tDec), "")
   dec = Val(tDec)
   Zbroji
End Sub

Sub Zbroji ()
   Dim s$, i%, c%, p%, red() As Double
   Dim suma#, temp$
   ReDim red(200)

   c = 0
   s = Trim(t)
   While Len(s) > 0
      c = c + 1
      p = InStr(s, Chr(13) & Chr$(10)): If p = 0 Then p = 200
      red(c) = Val(Left(s, p))
      s = Mid(s, p + 2)
   Wend


   'Cls
   'For i = 1 To c
   '   Print red(i)
   '   suma = suma + red(i)
   'Next
   t2 = ""
   
   For i = 1 To c
      temp = temp & Format(red(i), "#,##0." & String(dec, "0")) & Chr(13) & Chr$(10)
      suma = suma + red(i)
   Next
   t2 = temp
   lRez = Format(suma, "#,##0." & String(dec, "0"))
End Sub

