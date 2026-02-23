VERSION 2.00
Begin Form up_sort 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Sortiranje"
   ClientHeight    =   4200
   ClientLeft      =   2085
   ClientTop       =   2085
   ClientWidth     =   5475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   4605
   KeyPreview      =   -1  'True
   Left            =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5475
   Top             =   1740
   Width           =   5595
   Begin SSCommand CZaboravi 
      Caption         =   "&Odustani"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   3120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3660
      Width           =   1155
   End
   Begin SSCommand CSortiraj 
      Caption         =   "&Sortiraj"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   1200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3660
      Width           =   1155
   End
   Begin SSCommand Cukloni 
      Caption         =   "< Ukloni"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1155
   End
   Begin SSCommand Adddes 
      Caption         =   "Dodaj (-) >"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1500
      Width           =   1155
   End
   Begin SSCommand Addasc 
      Caption         =   "Dodaj (+) >"
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   1155
   End
   Begin ListBox List2 
      Height          =   3540
      Left            =   3360
      TabIndex        =   1
      Top             =   60
      Width           =   2055
   End
   Begin ListBox List1 
      Height          =   3540
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2055
   End
End
Option Explicit

Sub addasc_Click ()
Dim i As Integer, j As Integer
i = 0
j = -1
For i = 0 To list1.ListCount - 1
    If list1.Selected(i) Then j = i
Next i
If j <> -1 Then
    list2.AddItem "+" + list1.List(j)
    list2.ItemData(list2.NewIndex) = list1.ItemData(j)
    list1.RemoveItem j
End If
End Sub

Sub adddes_Click ()
Dim i As Integer, j As Integer
i = 0
j = -1
For i = 0 To list1.ListCount - 1
    If list1.Selected(i) Then j = i
Next i
If j <> -1 Then
    list2.AddItem "-" + list1.List(j)
    list2.ItemData(list2.NewIndex) = list1.ItemData(j)
    list1.RemoveItem j
End If
End Sub

Sub CSortiraj_Click ()
Dim i As Integer
   
If list2.ListCount > 0 Then
str1 = ""
For i = 0 To list2.ListCount - 1
    str1 = str1 & Left(list2.List(i), 1) & list2.ItemData(i) & ","
Next i
   str1 = Left(str1, Len(str1) - 1)
Else
   str1 = " "
End If
   
Unload Me
End Sub

Sub CUkloni_click ()
Dim i As Integer, j As Integer
i = 0
j = -1
For i = 0 To list2.ListCount - 1
    If list2.Selected(i) Then j = i
Next i
If j <> -1 Then
    list1.AddItem Right(list2.List(j), Len(list2.List(j)) - 1)
    list1.ItemData(list1.NewIndex) = list2.ItemData(j)
    list2.RemoveItem j
End If
End Sub

Sub CZaboravi_Click ()
str1 = ""
Unload Me
End Sub

Sub Form_KeyPress (keyascii As Integer)
If keyascii = 27 Then CZaboravi_Click
If keyascii = 13 Then CSortiraj_Click
If keyascii = 43 Then addasc_Click
If keyascii = 45 Then adddes_Click
End Sub

Sub Form_Load ()
centerform Me
End Sub

Sub List1_DblClick ()
addasc_Click
End Sub

Sub List2_DblClick ()
CUkloni_click
End Sub

