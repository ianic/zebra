Option Explicit

Sub dugmad (frm As Form, i%, e%, o%, f%, s%, n%, u%, b%, z%)
'izbor dugmadi koja se prikazuju te namještanje širine i visine forme
   frm.cizaberi.Visible = i * (-1)
   frm.cedit.Visible = e * (-1)

   frm.tg1.Editable = e * (-1)
   frm.tg1.MarqueeStyle = 3 - e

   frm.COsvjezi.Visible = o * (-1)
   frm.CFilter.Visible = f * (-1)
   frm.CSort.Visible = s * (-1)

   frm.cnovi.Visible = n * (-1)
   frm.Cupisi.Visible = u * (-1)
   frm.cbrisi.Visible = b * (-1)
   frm.cmdBrZapisa.Visible = z * (-1)
End Sub

Sub tg_combo (f As Form, col%, tbl$)
   f.tg1.VlistColumn = col
   Dim s As Snapshot, i%
   Dim b As Database
   Set b = OpenDatabase(f.Data1.DatabaseName)
   Set s = b.CreateSnapshot(tbl)
   i = 1
   Do Until s.EOF
      f.tg1.VlistText(i) = s.Fields(1).Value
      f.tg1.VlistData(i) = s.Fields(0).Value
      i = i + 1
      s.MoveNext
   Loop

   s.Close
End Sub

Sub tg_combo1 (f As Form, col%, tbl$)
   f.tg1.VlistColumn = col
   Dim s As Snapshot, i%
   Dim b As Database
   Set b = OpenDatabase(f.Data1.DatabaseName)
   Set s = b.CreateSnapshot(tbl)
   i = 1
   Do Until s.EOF
      f.tg1.VlistText(i) = s.Fields(0).Value
      f.tg1.VlistData(i) = s.Fields(0).Value
      i = i + 1
      s.MoveNext
   Loop
   s.Close
End Sub

Sub unifilter3 (f As Form, sn As Snapshot)

On Error GoTo filtererr
mp 11
Dim filt As New up_filt, i%, j%

For i = 0 To f.tg1.Columns - 1
If sn.Fields(f.tg1.ColumnField(i + 1)).Type <> 12 Then

   If i > 0 Then Load filt.Label1(i)
   filt.Label1(i).Top = 170 + 330 * i
   filt.Label1(i).Visible = True
   filt.Label1(i).Caption = f.tg1.ColumnName(i + 1)'sn.fields(f.tg1.columnfield(i + 1)).Name
f.tg1.VlistColumn = i + 1
j = f.tg1.VlistStyle And 1
If j = 1 Then
'combo box
      If i > 0 Then Load filt.Combo1(i)
      filt.Label1(i).Tag = ""
      filt.Combo1(i).Visible = True
      filt.Combo1(i).Top = 90 + 330 * i
      filt.Combo1(i).Tag = DajPunoIme(sn.Fields(f.tg1.ColumnField(i + 1)))
      filt.Combo1(i).AddItem ""
      For j = 1 To f.tg1.VlistCount
         filt.Combo1(i).AddItem f.tg1.VlistText(j)
         If IsNumeric(f.tg1.VlistData(j)) Then
            filt.Combo1(i).ItemData(filt.Combo1(i).NewIndex) = CLng(f.tg1.VlistData(j))
         Else
            filt.Label1(i).Tag = "T"
         End If
      Next j

Else
   Select Case sn.Fields(f.tg1.ColumnField(i + 1)).Type
   Case 1'boolean ,pravi combo sa Yes,No
      If i > 0 Then Load filt.Combo1(i)
      filt.Label1(i).Tag = ""
      filt.Combo1(i).Visible = True
      filt.Combo1(i).Top = 90 + 330 * i
      filt.Combo1(i).Tag = DajPunoIme(sn.Fields(f.tg1.ColumnField(i + 1)))
      filt.Combo1(i).AddItem ""
      filt.Combo1(i).AddItem "Da"
      filt.Combo1(i).ItemData(filt.Combo1(i).NewIndex) = -1
      filt.Combo1(i).AddItem "Ne"
      filt.Combo1(i).ItemData(filt.Combo1(i).NewIndex) = 0
      filt.Combo1(i).Width = 800
   Case 10'text
      If i > 0 Then Load filt.Text1(i)
      filt.Text1(i).Visible = True
      filt.Text1(i).Top = 90 + 330 * i
      filt.Text1(i).Tag = DajPunoIme(sn.Fields(f.tg1.ColumnField(i + 1)))
      filt.Label1(i).Tag = "1"
   Case 2, 3, 4'byte,integer,long
      If i > 0 Then Load filt.Text3(i)
      filt.Text3(i).Visible = True
      filt.Text3(i).Top = 90 + 330 * i
      filt.Text3(i).Tag = DajPunoIme(sn.Fields(f.tg1.ColumnField(i + 1)))  'sn.Fields(f(I)).sourcetable & "." & sn.Fields(f(I)).name
      filt.Label1(i).Tag = "3"
      
      If i > 0 Then Load filt.Text33(i)
      filt.Text33(i).Visible = True
      filt.Text33(i).Top = 90 + 330 * i
      filt.Text33(i).Tag = DajPunoIme(sn.Fields(f.tg1.ColumnField(i + 1)))  'sn.Fields(f(I)).sourcetable & "." & sn.Fields(f(I)).name
      'filt.Label1(I).Tag = "3"
   Case 6, 7'single,double
      If i > 0 Then Load filt.Text2(i)
      filt.Text2(i).Visible = True
      filt.Text2(i).Top = 90 + 330 * i
      filt.Text2(i).Tag = DajPunoIme(sn.Fields(f.tg1.ColumnField(i + 1)))
      filt.Label1(i).Tag = "2"
      
      If i > 0 Then Load filt.Text22(i)
      filt.Text22(i).Visible = True
      filt.Text22(i).Top = 90 + 330 * i
      filt.Text22(i).Tag = DajPunoIme(sn.Fields(f.tg1.ColumnField(i + 1)))
   Case 8'date
      If i > 0 Then Load filt.MaskedEdit1(i)
      filt.MaskedEdit1(i).Visible = True
      filt.MaskedEdit1(i).Top = 90 + 330 * i
      filt.MaskedEdit1(i).Tag = DajPunoIme(sn.Fields(f.tg1.ColumnField(i + 1)))
      
      If i > 0 Then Load filt.MaskedEdit2(i)
      filt.MaskedEdit2(i).Visible = True
      filt.MaskedEdit2(i).Top = 90 + 330 * i
      filt.MaskedEdit2(i).Tag = DajPunoIme(sn.Fields(f.tg1.ColumnField(i + 1)))
   End Select
End If
End If
Next i
filt.Height = 1300 + 330 * i
filt.cPrihvati.Top = 330 * i + 250
filt.cZaboravi.Top = 330 * i + 250

centerform filt
DoEvents
mp 0

filt.Show 1
Exit Sub
filtererr:
   str1 = ""
   Exit Sub
End Sub

Sub unisort3 (f As Form, sn As Snapshot) '(baza As Database, sel$, pom1$, ByVal pom$)

On Error GoTo sorterr
mp 11
Dim so As New up_sort, i%, ord$, pom$, i1%
so.List1.Clear
For i = 0 To f.tg1.Columns - 1
If sn.Fields(f.tg1.ColumnField(i + 1)).Type <> 12 Then
   so.List1.AddItem f.tg1.ColumnName(i + 1)
   so.List1.ItemData(so.List1.NewIndex) = sn.Fields(f.tg1.ColumnField(i + 1)).OrdinalPosition
End If
Next i

so.List2.Clear
mp 0
so.Show 1
If Len(str1) > 0 Then
If Len(str1) = 1 Then
   ord = " "
   Exit Sub
End If

ord = ""
pom = str1

Do
   i1 = InStr(pom, ",")
   If i1 > 0 Then
      ord = ord & DajPunoIme(sn.Fields(CInt(Mid(pom, 2, i1 - 2))))
      If Left(pom, 1) = "-" Then ord = ord & " desc "
      ord = ord & ","
      pom = Right(pom, Len(pom) - i1)
   Else
      ord = ord & DajPunoIme(sn.Fields(CInt(Right(pom, Len(pom) - 1))))
      If Left(pom, 1) = "-" Then ord = ord & " desc "
   End If
Loop Until i1 = 0
   str1 = ord
Else
   str1 = ""
End If
Exit Sub

sorterr:
   mp 0
   str1 = ""
   Exit Sub
End Sub

Sub urediTrueGrid (f As Form, tg1 As TrueGrid)
tg1.BackColor = &HC0C0C0
tg1.BorderStyle = 1
tg1.EditBackColor = &HFFFFFF
tg1.EditForeColor = &H0&
tg1.ExtendRightCol = False
tg1.SelectedBackColor = &H0&
tg1.SelectedForeColor = &H80FFFF
tg1.VertLines = 2
tg1.HorzLines = 2
tg1.VertColor = &H808080
tg1.HorzColor = &H808080
tg1.HeadBackColor = CLng(inidaj("Boje", "tablice", Str(8421376)))
tg1.HeadForeColor = &H80000009
End Sub

