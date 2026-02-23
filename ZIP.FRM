VERSION 2.00
Begin Form frmzip 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BACKUP"
   ClientHeight    =   2595
   ClientLeft      =   1365
   ClientTop       =   1770
   ClientWidth     =   4620
   ControlBox      =   0   'False
   Height          =   3000
   Icon            =   ZIP.FRX:0000
   Left            =   1305
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4620
   Top             =   1425
   Width           =   4740
   Begin SSFrame Frame3D2 
      Caption         =   "Mjesto arhiviranja"
      Height          =   1275
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   4515
      Begin OptionButton optB 
         BackColor       =   &H00C0C0C0&
         Caption         =   "sve baze u direktoriju"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   900
         Visible         =   0   'False
         Width           =   4095
      End
      Begin OptionButton optB 
         BackColor       =   &H00C0C0C0&
         Caption         =   "trenutna baza"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   660
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   4095
      End
      Begin TextBox Text1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   270
         Width           =   4365
      End
   End
   Begin CommandButton cmdOdustani 
      Caption         =   "&Odustani"
      Height          =   375
      Left            =   2730
      TabIndex        =   4
      Top             =   2160
      Width           =   1845
   End
   Begin SSPanel lblStatusBar 
      FloodColor      =   &H00808000&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   4515
   End
   Begin SSPanel p 
      FloodColor      =   &H00808000&
      FloodType       =   1  'Left To Right
      Height          =   345
      Left            =   60
      TabIndex        =   3
      Top             =   1770
      Width           =   4515
   End
   Begin CommandButton cmdPrihvati 
      Caption         =   "&Prihvati"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   2160
      Width           =   1875
   End
   Begin TextBox txtZIP 
      Height          =   285
      Left            =   1950
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2220
      Visible         =   0   'False
      Width           =   765
   End
End
Option Explicit

Sub cmdOdustani_Click ()
   Unload Me
End Sub

Sub cmdPrihvati_Click ()
   Dim Attr
   On Error Resume Next
   Attr = GetAttr(Text1)
   If Err = 53 Then
      MsgBox "Ne postoji taj direktorij!"
      Exit Sub
   End If
   
   cmdOdustani.Visible = False
   cmdPrihvati.Visible = False
 
   On Error GoTo errzip

   If Not zip(putbaza, IIf(Right(Text1, 1) = Chr$(92), Text1, Text1 & "\") & "RB" & Format(Date, "ddmmyy")) Then GoTo errzip

   Unload Me
   Exit Sub

errzip:
   MsgBox "Backup nije uspio!"
   cmdOdustani.Visible = True
   cmdPrihvati.Visible = True
   Resume Next
End Sub

Sub Form_Load ()
   CenterForm Me
   Dim i%
   i = addUNZIP_SetParentWindowHandle(Me.hWnd)
   i = addUNZIP_SetWindowHandle(txtZIP.hWnd)
   Text1 = Backup
End Sub

Sub txtZIP_Change ()
    Dim cAdditem As String
    Dim cAction As String
    Dim lSize As Long
    Debug.Print txtZIP.Text
    cAction = GetPiece((txtZIP.Text), "|", 2)
    Select Case cAction
        'Case "view"
        '    cAdditem = GetFileName((txtZIP.Text)) & Chr$(9)
        '    lSize = GetFileOriginalSize((txtZIP.Text))
        '    g_lSize = g_lSize + lSize
        '    cAdditem = cAdditem & Str$(lSize) & Chr$(9)
        '    cAdditem = cAdditem & Str$(GetFileCompressedSize((txtZIP.Text))) & Chr$(9)
        '    cAdditem = cAdditem & Str$(GetFileCompressionRatio((txtZIP.Text))) & "%" & Chr$(9)
        '    cAdditem = cAdditem & GetFilePath((txtZIP.Text))
        '    colArchive.AddItem cAdditem
        '    g_iCount = g_iCount + 1
        'Case "error"
        'Case "warning"
        Case Else
            'cAdditem = Format$(cAction, ">&&&&&&&&&&&") & " " & GetFileName((txtZIP.Text))
            'cAdditem = cAdditem & " - " & Str$(GetFileCompressionRatio((txtZIP.Text))) & "%"
            cAdditem = " " & UCase(GetFileName((txtZIP.Text)))
            'cAdditem = cAdditem & " - " & Str$(GetFileCompressionRatio((txtZIP.Text))) & "%"
            lblStatusBar.Caption = cAdditem

            'p.FloodPercent = GetFileCompressionRatio()
            p.FloodPercent = Val(GetPiece((txtZIP.Text), "|", 7))
    End Select
    DoEvents
End Sub

Sub unzip (sto$, kamo$, joker$, restoreS%)
   Dim i%
   p.FloodType = 1
   p.FloodPercent = 0
   i = addUNZIP_ArchiveName(sto)
   i = addUNZIP_Include(joker)
   i = addUNZIP_RestoreStructure(restoreS)
   i = addUNZIP_ExtractTo(kamo)
   i = addUNZIP()
   p.FloodPercent = 0
   p.FloodType = 0
   p.Caption = ""
   lblStatusBar.Caption = ""
End Sub

Function zip (sto$, ime$)
zip = True
   Dim i%, jj%
jj = 0
josjednom:
jj = jj + 1
   p.FloodType = 1
   p.FloodPercent = 0
   i = addZIP_SetWindowHandle(txtZIP.hWnd)
   i = addZIP_Include(sto)
   i = addZIP_ArchiveName(ime)
   i = addZIP_Overwrite(11)
   i = addZIP_SaveStructure(1)
   Dim putt$
   putt = sto
   On Error Resume Next
      Kill Left(putt, Len(putt) - 3) & "ldb"
   On Error GoTo 0
   DoEvents
   i = addZIP()
   If i = 0 And jj = 1 Then GoTo josjednom

   If i = 0 Then
      MsgBox "Baza " & sto & " je zauzeta!." & Chr$(13) & "Zatvorite aplikacije i pokušajte ponovno!"
      zip = False
   Else
      MsgBox "Backup uspješan! Kreirana datoteka " & ime & ".zip!"
   End If
   p.FloodPercent = 0
   p.FloodType = 0
   p.Caption = ""
   lblStatusBar.Caption = ""
End Function

