VERSION 2.00
Begin Form up_filt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Filter"
   ClientHeight    =   2460
   ClientLeft      =   3045
   ClientTop       =   3510
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   2970
   KeyPreview      =   -1  'True
   Left            =   2985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6240
   Top             =   3060
   Width           =   6360
   Begin CommandButton cZaboravi 
      BackColor       =   &H00404040&
      Caption         =   "&Odustani"
      Height          =   465
      Left            =   3450
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1155
   End
   Begin CommandButton cPrihvati 
      BackColor       =   &H00000080&
      Caption         =   "&Prihvati"
      Default         =   -1  'True
      Height          =   465
      Left            =   1740
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1155
   End
   Begin TextBox Text33 
      Height          =   315
      Index           =   0
      Left            =   3750
      TabIndex        =   8
      Top             =   750
      Visible         =   0   'False
      Width           =   1905
   End
   Begin TextBox Text22 
      Height          =   315
      Index           =   0
      Left            =   3750
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MaskEdBox MaskedEdit2 
      Height          =   315
      Index           =   0
      Left            =   3030
      Mask            =   "##/##/####"
      MaxLength       =   10
      PromptChar      =   "_"
      TabIndex        =   6
      Top             =   1410
      Visible         =   0   'False
      Width           =   1275
   End
   Begin ComboBox Combo1 
      Height          =   300
      Index           =   0
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   2985
   End
   Begin MaskEdBox MaskedEdit1 
      Height          =   315
      Index           =   0
      Left            =   1620
      Mask            =   "##/##/####"
      MaxLength       =   10
      PromptChar      =   "_"
      TabIndex        =   2
      Top             =   1410
      Visible         =   0   'False
      Width           =   1275
   End
   Begin TextBox Text3 
      Height          =   315
      Index           =   0
      Left            =   1620
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   1905
   End
   Begin TextBox Text2 
      Height          =   315
      Index           =   0
      Left            =   1620
      TabIndex        =   4
      Top             =   420
      Visible         =   0   'False
      Width           =   1905
   End
   Begin TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   2955
   End
   Begin Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   990
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   585
   End
End
Option Explicit

Sub CPrihvati_Click ()
On Error Resume Next
DoEvents
Dim i%, n%
Dim tmp
str1 = " "
For i = 0 To controls.Count - 1
If controls(i).Visible Then

   If TypeOf controls(i) Is TextBox Then
   If Len(controls(i).Text) > 0 Then
   n = controls(i).Index
   Select Case label1(n).Tag
   Case "1" 'string
      str1 = str1 & text1(n).Tag & " like '" & text1(n).Text & "*'  and "
   Case "2"'broj DOUBLE
      'str1 = str1 & text2(n).Tag & " = " & text2(n).Text & " and "
      str1 = str1 & text2(n).Tag & " between " & Str(CDbl(text2(n).Text)) & " and " & Str(CDbl(text22(n).Text)) & " and "
      label1(n).Tag = ""
   Case "3"'broj INT
      'str1 = str1 & text3(n).Tag & " = " & text3(n).Text & " and "
      str1 = str1 & text3(n).Tag & " between " & Str(Int(text3(n).Text)) & " and " & Str(Int(text33(n).Text)) & " and "
      label1(n).Tag = ""
   End Select
   End If
   End If

   If TypeOf controls(i) Is MaskEdBox Then  'DATUM
   If IsDate(controls(i).Text) Then
      n = controls(i).Index
      str1 = str1 & MaskedEdit1(n).Tag & " between " & NapraviDatum2((MaskedEdit1(n).Text)) & " and " & NapraviDatum2((MaskedEdit2(n).Text)) & " and "
      
      'str1 = str1 & maskededit1(n).Tag & " = #" + Format(CVDate(maskededit1(n).Text), "mm") + "/" + Format(CVDate(maskededit1(n).Text), "dd") + "/" + Format(CVDate(maskededit1(n).Text), "yy") + "# and "
   End If
   End If
   
   If TypeOf controls(i) Is ComboBox Then
   If controls(i).ListIndex > 0 Then
      n = controls(i).Index
      If label1(n).Tag = "T" Then
         str1 = str1 & combo1(n).Tag & " = '" & combo1(n).List(combo1(n).ListIndex) & "' and "
      Else
         str1 = str1 & combo1(n).Tag & " = " & combo1(n).ItemData(combo1(n).ListIndex) & " and "
      End If
   End If
   End If

End If
Next i
str1 = Left(str1, Len(str1) - 5)
Unload Me
End Sub

Sub CZaboravi_Click ()
str1 = ""
Unload Me
End Sub

Sub Form_KeyPress (KeyAscii As Integer)

If KeyAscii = 27 Then CZaboravi_Click
'If KeyAscii = 13 Then DoEvents: cprihvati_click
End Sub

Sub MaskedEdit1_KeyPress (Index As Integer, KeyAscii As Integer)
datkp MaskedEdit1(Index), KeyAscii
End Sub

Sub MaskedEdit1_KeyUp (Index As Integer, KeyCode As Integer, Shift As Integer)
   MaskedEdit2(Index).Text = MaskedEdit1(Index).Text
End Sub

Sub MaskedEdit1_LostFocus (Index As Integer)
datlf MaskedEdit1(Index)
End Sub

Sub MaskedEdit2_KeyPress (Index As Integer, KeyAscii As Integer)
datkp MaskedEdit2(Index), KeyAscii
End Sub

Sub MaskedEdit2_LostFocus (Index As Integer)
datlf MaskedEdit2(Index)
End Sub

Sub Text2_KeyPress (Index As Integer, KeyAscii As Integer)
DoEvents
KeyAscii = dblkp(KeyAscii)
End Sub

Sub Text2_KeyUp (Index As Integer, KeyCode As Integer, Shift As Integer)
   text22(Index) = text2(Index)

End Sub

Sub Text2_LostFocus (Index As Integer)
dbllf text2(Index)
dbllf text22(Index)
End Sub

Sub Text22_KeyPress (Index As Integer, KeyAscii As Integer)
KeyAscii = dblkp(KeyAscii)
End Sub

Sub Text22_LostFocus (Index As Integer)
dbllf text22(Index)
End Sub

Sub Text3_KeyUp (Index As Integer, KeyCode As Integer, Shift As Integer)
   text33(Index) = text3(Index)

End Sub

Sub Text3_LostFocus (Index As Integer)
intlf text3(Index)
intlf text33(Index)
text33(Index) = text3(Index)
DoEvents
End Sub

Sub Text33_LostFocus (Index As Integer)
intlf text33(Index)
End Sub

