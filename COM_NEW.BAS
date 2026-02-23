Option Explicit
Global dbl1 As Double

Sub cboNplac (c As ComboBox)
    c.Clear
    c.AddItem "bezgotovinsko plaæanje"
    c.AddItem "gotovinsko plaæanje"
    c.AddItem "IZVOZ"
    c.AddItem "ne podliježe PDV-u OSTALO"
    c.ListIndex = 0
End Sub

Sub cboOznaka (c As ComboBox)
    c.Clear
    c.AddItem "T - tiskara"
    c.AddItem "K - knjižara"
    c.ListIndex = 0
End Sub

Sub cboOZNAKA1 (O$, c As ComboBox)
Select Case O
Case "T"
    c.ListIndex = 0
Case "K"
    c.ListIndex = 1
End Select
End Sub

Sub cbotiprac (c As ComboBox)
    c.Clear
    c.AddItem "R-1"
    c.AddItem "R-2"
    c.ListIndex = 0
End Sub

Sub cbotiprac1 (t$, c As ComboBox)
Select Case t
Case "R-1"
    c.ListIndex = 0
Case "R-2"
    c.ListIndex = 1
Case Else
    c.ListIndex = 0
End Select
End Sub

Function cCDbl (c) As Double
If IsNumeric(c) Then
    cCDbl = CDbl(c)
Else
    cCDbl = 0
End If
End Function

Sub custgrid (data1 As Data, tg1 As TrueGrid, g1 As Grid, value%)
If value Then
Dim i%, s As snapshot, j%, ps$, ok%
Set s = data1.Recordset.CreateSnapshot()
j = 0
g1.Row = 0
For i = 1 To tg1.Columns

If s.Fields(tg1.ColumnField(i)).Type = 7 Then
ok = True
If LCase(Left(tg1.ColumnField(i), 1)) = "c" Then ok = False
If LCase(tg1.ColumnField(i)) = "rabp" Then ok = False
If LCase(tg1.ColumnField(i)) = "pdvp" Then ok = False
If ok Then 'preskoèi cijena i postotke
    g1.Cols = j + 1
    g1.Col = j
    g1.Text = tg1.ColumnName(i)
    g1.ColWidth(j) = 1200'tg1.ColumnRectWidth(i) + 20
    g1.ColAlignment(j) = 1
    j = j + 1
End If
End If
Next i
g1.Width = j * 1200 + j * 20
RacSumu data1, tg1, g1
g1.Visible = True
Else
    g1.Visible = False
End If
End Sub

Function DajBrTxt (s As String) As Double
On Error Resume Next
DajBrTxt = 0
If IsNumeric(s) Then
    DajBrTxt = CDbl(s)
End If
End Function

Function db (f As Field)
db = IIf(IsNull(f), 0, f)
End Function

Function dbf (f As Field)
dbf = Format(db(f), "#,##0.00#")
End Function

Function dd (f As Field)
If IsNull(f) Then
    dd = ""
Else
    dd = Format(f, "dd.mm.yyyy")
End If
End Function

Function ds (f As Field)
ds = IIf(IsNull(f), "", f)
End Function

Sub greska ()
MsgBox Error, 16, "Greška " & Err
End Sub

Sub RacSumu (data1 As Data, tg1 As TrueGrid, g1 As Grid)
If data1.Recordset.RecordCount = 0 Then Exit Sub
mp 11
h "Raèunam sumu"
Dim i%, c() As String, j%, k%
g1.Row = 0
k = 0
For i = 0 To g1.Cols - 1
    g1.Col = i
    For j = 1 To tg1.Columns
        If tg1.ColumnName(j) = g1.Text Then
            ReDim Preserve c(k)
            c(k) = tg1.ColumnField(j)
            k = k + 1
            GoTo next1
        End If
    Next j
next1:
Next i

If k = 0 Then Exit Sub
k = k - 1
Dim s As snapshot, d() As Double
ReDim d(0 To k)
Set s = data1.Recordset.CreateSnapshot()
s.MoveFirst
Do
    For i = 0 To k
        d(i) = d(i) + db(s(c(i)))
    Next i
    s.MoveNext
Loop Until s.EOF

g1.Row = 1
For i = 0 To g1.Cols - 1
    g1.Col = i
    g1.Text = Format(d(i), "Standard")
Next i
h ""
mp 0
End Sub

Sub sel ()
   Dim c As Control
   On Error Resume Next
   Set c = screen.ActiveControl
   
   c.SelStart = 0
   c.SelLength = Len(c)
   
End Sub

Sub zaok (d As Double, dp%)
Dim s$
On Error Resume Next
s = "0." & String(dp, "0")
d = Format(d, s)
End Sub

Function zaokF (d As Double, dp%) As Double
Dim s$
On Error Resume Next
s = "0." & String(dp, "0")
zaokF = Format(d, s)
End Function

