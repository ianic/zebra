VERSION 2.00
Begin Form baze 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Izbor baze podataka"
   ClientHeight    =   2130
   ClientLeft      =   3690
   ClientTop       =   1575
   ClientWidth     =   3930
   Height          =   2535
   Icon            =   BAZE.FRX:0000
   KeyPreview      =   -1  'True
   Left            =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   3930
   Top             =   1230
   Width           =   4050
   Begin FileListBox File1 
      Height          =   1980
      Left            =   150
      Pattern         =   "*.mdb"
      TabIndex        =   2
      Top             =   60
      Width           =   1875
   End
   Begin SSCommand cOdustani 
      Caption         =   "&Odustani"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   2520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   690
      Width           =   1155
   End
   Begin SSCommand ok 
      Caption         =   "&Izaberi"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   2520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1155
   End
End
Option Explicit
Dim pb$

Sub COdustani_Click ()
'zatvorisve
Unload Me
End Sub

Sub File1_DblClick ()
ok_Click
End Sub

Sub Form_KeyPress (keyascii As Integer)
If keyascii = 27 Then COdustani_Click
If keyascii = 13 Then ok_Click

End Sub

Sub Form_Load ()
centerform Me
'inifile = app.Path & app.EXEName & ".ini"
file1.Path = inidaj("baze", "PutDoBaze", (app.Path) & "\baze")

Dim s$, i%
s = inidaj("baze", "ZadnjaBaza", "")
If file1.ListCount > 0 Then file1.ListIndex = file1.ListCount - 1
For i = 0 To file1.ListCount - 1
   If file1.List(i) = s Then file1.ListIndex = i: Exit For
Next i
End Sub

Sub List1_DblClick ()
ok_Click
End Sub

Sub List1_KeyPress (keyascii As Integer)
If keyascii = 13 Then ok_Click
End Sub

Sub ok_Click ()
putbaza = file1.Path & "\" & file1.List(file1.ListIndex)
putbaza1 = file1.Path
frm_mdi.Show
frm_mdi.Caption = frm_mdi.Caption & " [" & file1.List(file1.ListIndex) & "]"
Dim b As database
Set b = OpenDatabase(putbaza)
Dim i%, ok%
For i = 0 To b.TableDefs.Count - 1
   If b.TableDefs(i).Name = "ul_faktura" Then ok = True
Next i
If Not ok Then frm_mdi.mnu_5.Visible = False
inipisi "Baze", "ZadnjaBaza", file1.List(file1.ListIndex)
Unload Me
'pb = putbaza & list1.List(list1.ListIndex)
'On Error GoTo kriva
'Dim s As snapshot, baza As database
'If Not bazaok(pb) Then GoTo kriva
'Set baza = OpenDatabase(pb)
'Set s = baza.CreateSnapshot("OsnovniPodaci")
's.Close
'baza.Close
'putbaza = pb
'par1 = 1
'Unload Me
'logo.Show
'logo.ZOrder 0
'DoEvents
'Exit Sub
'kriva:
'    mp 0
'   hrmb "Neispravna baza !!!", 16, "Greška"
'   Exit Sub
End Sub

