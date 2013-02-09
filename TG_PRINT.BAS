Option Explicit
Global TgPrintBaz$
Dim tg_suma_g1%

Function DajPolje (polja$, i%) As Integer
On Error GoTo dajpoljeerr
DajPolje = False
If Mid(polja, i, 1) = "1" Then DajPolje = True' Else DajPolje = True

Exit Function
dajpoljeerr:
    DajPolje = False
    Exit Function
End Function

Function DajSirina (sirina$, i%) As Integer
On Error GoTo dajsirinaerr

DajSirina = Format(Mid(sirina, i * 2 - 1, 2), "00")

Exit Function
dajsirinaerr:
    DajSirina = -1
    Exit Function
End Function

'sir-širina kolone
'head-heading
'Vraæa širinu kolone pri grafièkom ispisu tako da stane heading
Function DajSirinuKolone (sir%, head$) As Long
   DajSirinuKolone = sir * 100
   'Dim p1&, p2&
   'p1 = sir * printer.TextWidth("0")'po postavci
   'p2 = printer.TextWidth(head) + printer.TextWidth("0")'da stane heder
   'DajSirinuKolone = IIf(p1 > p2, p1, p2)
End Function

Function imazapis (m%, id$) As Integer
imazapis = False
On Error GoTo izerr
If Len(Dir$(TgPrintBaz)) > 0 Then
    Dim s As snapshot, b As database
    Set b = OpenDatabase(TgPrintBaz)
    Set s = b.CreateSnapshot("select * from tg_print where m=" & m)
    If Len(id) > 0 Then s.FindFirst "id='" & id & "'"
    If s.RecordCount > 0 Then
        imazapis = True
        Exit Function
    End If
End If
Exit Function
izerr:
   Exit Function
End Function

Sub punitgpr (f As Form, t As TrueGrid, d As Data)
Dim i%, maxheight&, curtop&, Col%, plusleft%, j%, pola&
maxheight = screen.Height * .65
Col = 0
j = 0
plusleft = 0
pola = Int((t.Columns + 1) / 2)
On Error Resume Next
tgpr.Text1 = f.Caption
For i = 1 To t.Columns
   If i > 1 Then
      Load tgpr.check(i)
      Load tgpr.chksum(i)
      Load tgpr.mask(i)
      tgpr.chksum(i).ZOrder 0
      tgpr.mask(i).ZOrder 0
      tgpr.check(i).Visible = True
      tgpr.mask(i).Visible = True
      curtop = tgpr.check(1).Top + 270 * (i - 1 - j)
    If i = pola + 1 Then
      j = i - 1
      plusleft = plusleft + 3100
      curtop = tgpr.check(1).Top
    End If
      tgpr.check(i).Top = curtop
      tgpr.mask(i).Top = curtop + 30
      tgpr.chksum(i).Top = curtop
      tgpr.chksum(i).Left = tgpr.chksum(i).Left + plusleft
      tgpr.mask(i).Left = tgpr.mask(i).Left + plusleft
      tgpr.check(i).Left = tgpr.check(i).Left + plusleft
   End If
      tgpr.chksum(i).Visible = False
      tgpr.chksum(i) = False
      Select Case d.Recordset.Fields(t.ColumnField(i)).Type
      Case 10'text
        tgpr.check(i).ForeColor = &H808000
      Case 2, 3, 4'byte,integer,long
        tgpr.chksum(i).Visible = True
        tgpr.chksum(i) = False
        tgpr.check(i).ForeColor = &H800000
      Case 6, 7'single,double
        tgpr.chksum(i).Visible = True
        tgpr.chksum(i) = True
        tgpr.check(i).ForeColor = &H80&
      Case 8'date
      End Select
   tgpr.check(i).Caption = t.ColumnName(i)
   tgpr.mask(i).Text = t.ColumnWidth(i)
Next i

tgpr.Height = tgpr.check(1).Top + 280 * pola + 1420
tgpr.pnl3.Top = tgpr.Height - 1500
tgpr.cmd3.Top = tgpr.Height - 1150
tgpr.cmd4.Top = tgpr.Height - 1150
tgpr.panf1.Top = tgpr.Height - 1150
tgpr.panf2.Top = tgpr.Height - 850
tgpr.cmd1.Top = tgpr.Height - 1150
tgpr.cmd2.Top = tgpr.Height - 1150

CenterForm tgpr
'tgpr.cmd3.Top = tgpr.check(i - 1).Top + 500
'tgpr.cmd4.Top = tgpr.cmd3.Top
'tgpr.panf1.Top = tgpr.cmd4.Top
'tgpr.panf2.Top = tgpr.cmd4.Top
'tgpr.cmd1.Top = tgpr.cmd4.Top + 500
'tgpr.cmd2.Top = tgpr.cmd1.Top
'tgpr.height = tgpr.cmd1.Top + 900

tgpr.panf1.Caption = printer.FontName
tgpr.panf2.Caption = Format(printer.FontSize, "Fixed")
If draftispis Then
    tgpr.panf1.Visible = False
    tgpr.panf2.Visible = False
    tgpr.cmd4.Visible = False
    tgpr.cmd3.Visible = False
End If

End Sub

Sub punitgprizbaze (ByVal m%, ByVal id$)
On Error GoTo piberr
Dim i%
Dim polj$, sirina$, sir%, suma$
'ptonaði zapis u bazi

    Dim s As snapshot, b As database
    Set b = OpenDatabase(TgPrintBaz)
    Set s = b.CreateSnapshot("select * from tg_print where m=" & m)
    If Len(id) > 0 Then
        s.FindFirst "id='" & id & "'"
    Else
        'napuni combo u tgpr
        Do Until s.EOF
            tgpr.Combo1.AddItem s!id
        s.MoveNext
        Loop
        tgpr.Combo1.ListIndex = 0
        s.MoveFirst
        tgpr.Combo1.Tag = m
    End If

'napuni tgpr
tgpr.Text1.Text = ds(s!zaglavlje)
tgpr.Text2.Text = ds(s!zaglavlje1)
polj = ds(s!polja)
sirina = ds(s!sirina)
suma = ds(s!suma)
For i = 1 To Len(polj)
   tgpr.check(i).Value = DajPolje(polj, i)
   tgpr.chksum(i).Value = DajPolje(suma, i)
   sir = DajSirina(sirina, i)
   tgpr.mask(i) = Format(sir, "###")
Next i

If Not draftispis Then
    'postavi orijentaciju papira
    Load frmprnodabir
    frmprnodabir.optOrj(0) = s!PORT
    frmprnodabir.cmdPrihvati = True
    'frmprnodabir.cboPrn = "HP LaserJet 4/4M"
    Unload frmprnodabir
    'font name i size
    printer.EndDoc
    tgpr.panf1.Caption = s!fn
    tgpr.panf2.Caption = Format(s!fs, "Fixed")
    printer.FontSize = CDbl(db(s!fs))
    printer.FontName = (s!fn)
End If
Exit Sub
s.Close
b.Close
piberr:
   Debug.Print "Punitrprizbaze greška: " & Error
   Exit Sub
End Sub

'm- odreðuje mod ispisa, svaka konfiguracija trugrida treba imati razlièit
Sub tg_print (f As Form, t As TrueGrid, d As Data, m%)
If draftispis Then tg_print1 f, t, d, m: Exit Sub
'Dim mm%
'mm = m
mp 11
'On Error GoTo tg_print_err
Dim i%

Load tgpr
tgpr.Combo1.Clear

punitgpr f, t, d'napuni formu za podešavanje s podacima iz truegrid-a
If imazapis(m, "") Then punitgprizbaze m, ""

'prikaz forme za podešavanje ispisa
CenterForm tgpr
tgpr.Show 1
If Len(tgpr.Tag) > 0 Then 'GoTo tg_print_err 'izlazak-odabrano odustani
    printer.EndDoc
    Unload tgpr
    mp 0
    Exit Sub
End If
    
Dim fn$, fs$
fn = tgpr.panf1.Caption'font name
fs = tgpr.panf2.Caption'font size

'----------------------------
If tgpr.snimi.Value And Len(tgpr.Combo1.Text) > 0 Then upisizapis m, Left(tgpr.Combo1.Text, 30), t.Columns
mp 11
On Error GoTo tg_print_err

'priprema za ispis
Dim polja() As String, j%, xpoz&(), p&, head()  As String, tgpoz&(), suma#()
'Dim p1&, p2&
j = 0
p = 0
For i = 1 To t.Columns
If tgpr.check(i) Then
   ReDim Preserve polja(j)
   ReDim Preserve head(j)
   ReDim Preserve xpoz(j + 1)
   ReDim Preserve tgpoz(j + 1)
   ReDim Preserve suma(j)
   polja(j) = t.ColumnField(i)
   head(j) = t.ColumnName(i)
   tgpoz(j) = i
   If tgpr.chksum(i) Then suma(j) = 0 Else suma(j) = -1
   'odreðivanje velièine kolona
   'p1 = tgpr.mask(i) * printer.TextWidth("0")'po postavci
   'p2 = printer.TextWidth(head(j)) + printer.TextWidth("0")'da stane heder
   'p = p + IIf(p1 > p2, p1, p2)
   p = p + DajSirinuKolone((tgpr.mask(i)), head(j))
   xpoz(j + 1) = p
   j = j + 1
End If
Next i
If xpoz(j) > printer.ScaleWidth Then
    screen.MousePointer = 0
    If MsgBox("Ispis preko rubova papira! Želite li da nastavim s ispisom?", 36, "Upozorenje") = 7 Then
        Unload tgpr
        Exit Sub
    End If
    screen.MousePointer = 11
End If
'ispis zaglavlja
printer.FontName = fn
printer.FontSize = CDbl(fs) * 1.5
printer.Print tgpr.Text1.Text
printer.Print tgpr.Text2.Text': printer.Print
printer.FontSize = fs
If tgpr.di Then
    printer.Print "Datum ispisa: " & Format(Now, "dd.mm.yyyy"): printer.Print
End If


Dim s As snapshot, strana%
Set s = d.Recordset.CreateSnapshot()'kreiranje snapshota
Dim k%, y%
Dim st As String
strana = 1


ispis_naziva_polja:
printer.FontUnderline = True
'ispis naziva polja
For i = 0 To j - 1
   st = head(i)
   printer.CurrentX = xpoz(i)
   If s.Fields(polja(i)).Type = 7 Then
        printer.CurrentX = xpoz(i + 1) - printer.TextWidth(st)
   End If
   If i > 0 Then
        If s.Fields(polja(i - 1)).Type = 7 And s.Fields(polja(i)).Type <> 7 Then printer.CurrentX = xpoz(i) + printer.TextWidth("0")
   End If
   printer.Print st;
Next i
printer.Print
printer.FontUnderline = False
'ispis polja iz baze

Do Until s.EOF
DoEvents
For i = 0 To j - 1
   printer.CurrentX = xpoz(i)
   st = ""
        If IsNull(s.Fields(polja(i)).Value) Then
         st = ""
        Else
         If s.Fields(polja(i)).Type = 7 Then st = Format(s.Fields(polja(i)).Value, "Standard") Else st = s.Fields(polja(i)).Value
         'za logièka polja DA-NE
         If s.Fields(polja(i)).Type = 1 Then
            If s.Fields(polja(i)).Value = 0 Then st = "Ne" Else st = "Da"
         End If

        End If
   'rezanje predugaèkog teksta
   Do Until printer.TextWidth(st) < (xpoz(i + 1) - xpoz(i))
      st = Left(st, Len(st) - 1)
   Loop
   'On Error Resume Next
   If s.Fields(polja(i)).Type = 7 Then
    printer.CurrentX = xpoz(i + 1) - printer.TextWidth(st)
   End If
   If i > 0 Then
    If s.Fields(polja(i - 1)).Type = 7 And s.Fields(polja(i)).Type <> 7 Then printer.CurrentX = xpoz(i) + printer.TextWidth("0")
   End If
   printer.Print st;
   'sumiraj
    If suma(i) <> -1 Then
        On Error Resume Next
        suma(i) = suma(i) + db(s.Fields(polja(i)))
        On Error GoTo tg_print_err
    End If
Next i
printer.Print
s.MoveNext
If (printer.CurrentY > printer.ScaleHeight - 2 * printer.TextHeight("stranica " & strana & "   ")) Then
    printer.CurrentX = printer.ScaleWidth - printer.TextWidth("stranica " & strana & "   ")
    printer.CurrentY = printer.ScaleHeight - printer.TextHeight("stranica " & strana & "   ")
    printer.Print "stranica " & strana;
    strana = strana + 1
   ' If s.EOF Then
   '     h ""
   ' Else
        h "Stranica: " & strana
        printer.NewPage
        GoTo ispis_naziva_polja
   ' End If
End If
Loop

Dim x%
If tg_suma_g1 = 0 Then GoTo tg1_print_1'preskoèi sumu iz g1
On Error GoTo tg1_print_1
If f.g1.Visible Then
On Error Resume Next
printer.FontBold = True
    For i = 0 To j - 1
        For x = 0 To f.g1.Cols - 1
            f.g1.Row = 0
            f.g1.Col = x
            If f.g1.Text = head(i) Then
                f.g1.Row = 1
                st = f.g1.Text
                printer.CurrentX = xpoz(i + 1) - printer.TextWidth(st)
                printer.Print st;
                GoTo totali
            End If
        Next x
totali:
    Next i
End If

tg1_print_1:
printer.FontUnderline = True
For i = 0 To j - 1
    If suma(i) <> -1 Then
        st = Format(suma(i), "Standard")
        printer.CurrentX = xpoz(i + 1) - printer.TextWidth(st)
        printer.Print st;
    End If
Next i
printer.FontUnderline = False
If strana > 1 Then
    printer.CurrentX = printer.ScaleWidth - printer.TextWidth("stranica " & strana & "   ")
    printer.CurrentY = printer.ScaleHeight - printer.TextHeight("stranica " & strana & "   ")
    printer.Print "stranica " & strana;
End If
On Error GoTo tg_print_err
DoEvents
printer.EndDoc
Unload tgpr
mp 0
h ""
Exit Sub
tg_print_err:
'MsgBox Error, 16, "Greška"
On Error Resume Next
printer.EndDoc
Unload tgpr
h ""
mp 0
Exit Sub
End Sub

'ispis na matriène pisaèe iz TrueGrida
'm- odreðuje mod ispisa, svaka konfiguracija trugrida treba imati razlièit
Sub tg_print1 (f As Form, t As TrueGrid, d As Data, m%)

mp 11
'On Error GoTo tg_print_err1
Dim i%

Dim x1%, x2%, x3%, x%

m = m + draftispis
Load tgpr
tgpr.Combo1.Clear
punitgpr f, t, d'napuni formu za podešavanje s podacima iz truegrid-a
If imazapis(m, "") Then punitgprizbaze m, ""', PutDoBaze

'prikaz forme za podešavanje ispisa
tgpr.Show 1
If Len(tgpr.Tag) > 0 Then GoTo tg_print_err1 'izlazak-odabrano odustani

If tgpr.snimi.Value Then upisizapis m, Left(tgpr.Combo1.Text, 30), t.Columns
mp 11
'On Error GoTo tg_print_err1

'priprema za ispis
Dim polja() As String, j%, xpoz&(), p&, head() As String, suma#()
Dim He%
Dim sir%
Dim stil() As Integer
Dim MaxRed%, strana%, Red%

MaxRed = 62: strana = 1: Red = 0

j = 0
p = 0
For i = 1 To t.Columns
If tgpr.check(i) Then
   ReDim Preserve polja(j)
   ReDim Preserve head(j)
   ReDim Preserve br(j)
   ReDim Preserve xpoz(j + 1)
   ReDim Preserve stil(j)
   ReDim Preserve tip(j)
   ReDim Preserve suma(j)
   polja(j) = t.ColumnField(i)
   head(j) = t.ColumnName(i)
   If tgpr.chksum(i) Then suma(j) = 0 Else suma(j) = -1
   stil(j) = t.ColumnStyle(i)
   tip(j) = t.EditMask(i)
   br(j) = tgpr.mask(i)
   p = p + tgpr.mask(i) + 2
   sir = sir + tgpr.mask(i)
   xpoz(j + 1) = p
   j = j + 1
End If
Next i

If xpoz(j) > 160 Then 'maaximalno 160 znakova pri 20 cpi
    If MsgBox("Ispis preko rubova papira! Maksimalno je dopušteno 160 znakova u redu, a Vi imate " & xpoz(j) & "!!!" & Chr(13) & "Želite li da nastavim s ispisom?", 36, "Upozorenje") = 7 Then
        Unload tgpr
        mp 0
        Exit Sub
    End If
End If

'ispis zaglavlja
otvoriprn2 72
If mode = 0 Then Exit Sub
lq 1
    cpi 6
    bold 1
    Iob (tgpr.Text1.Text), 45, 2: NLine: Red = Red + 1
    Iob (tgpr.Text2.Text), 45, 2: NLine: Red = Red + 1
    NLine
    Red = Red + 1
lq 0
If tgpr.di Then 'datum ispisa
    cpi 17: bold 0
    Iob "Datum ispisa: ", 15, 10
    Iob Format(Now, "dd.MM.yyyy"), 15, 0: NLine: Red = Red + 1
End If

    cpi 20
    bold 0
    Iob "=", 160, 5
    NLine
    Red = Red + 1
    cpi 20
    For i = 0 To j - 1
        x = stil(i) And 2
        If x = 0 Then
            Iob Mid(head(i), 1, br(i) + 1), br(i) + 2, 10
        Else
            Iob Mid(head(i), 1, br(i) + 1) + " ", br(i) + 2, 1
        End If
    Next i
    NLine
    Red = Red + 1
    cpi 20
    bold 0
    Iob "-", 160, 5
    NLine
    Red = Red + 1

'ispis polja
Dim s As snapshot
Set s = d.Recordset.CreateSnapshot()
Dim k%, y%
Dim st As String
x = 0

If s.RecordCount = 0 Then Exit Sub
Do Until s.EOF
DoEvents

For i = 0 To j - 1
   st = ""
         If s.Fields(polja(i)).Type = 1 Then
            If s.Fields(polja(i)).Value = 0 Then st = "Ne" Else st = "Da"
         ElseIf s.Fields(polja(i)).Type = 8 Then
            st = dd(s.Fields(polja(i)))
         Else
            st = IIf(IsNull(s.Fields(polja(i)).Value), "", s.Fields(polja(i)).Value)
         End If
        x = stil(i) And 2
        If x = 0 Then
            Iob Mid(st, 1, br(i) + 1), br(i) + 2, 10   'left ustify
        Else
            Iob Mid(Format(st, tip(i)), 1, br(i) + 1) & " ", br(i) + 2, 1'right justify
        End If
    If suma(i) <> -1 Then
        suma(i) = suma(i) + db(s.Fields(polja(i)))
    End If
Next i

NLine
Red = Red + 1
s.MoveNext
'nova strana
If (Red >= MaxRed) And (Not s.EOF) Then
    NLine
    cpi 20
    bold 0
    Iob "=", 160, 5
    NLine
    Red = Red + 1
    
    NovaStrana strana
    Red = 1
    
    cpi 20
    bold 0
    Iob "=", 160, 5
    NLine
    Red = Red + 1
    cpi 20
    For i = 0 To j - 1
        x = stil(i) And 2
        If x = 0 Then
            Iob Mid(head(i), 1, br(i) + 1), br(i) + 2, 10
        Else
            Iob Mid(head(i), 1, br(i) + 1) + " ", br(i) + 2, 1
        End If
    Next i
    NLine
    Red = Red + 1
    cpi 20
    bold 0
    Iob "-", 160, 5
    NLine
    Red = Red + 1
End If
x = 0
Loop
'---------------------
DoEvents

    
    cpi 20
    bold 0
    Iob "=", 160, 5
    NLine

If tg_suma_g1 = 0 Then GoTo tg1_print1_1'preskoèi sumu iz g1
'ispis totala iz grida
On Error Resume Next
If f.g1.Visible Then
    For i = 0 To j - 1
        For x = 0 To f.g1.Cols - 1
            f.g1.Row = 0
            f.g1.Col = x
            If f.g1.Text = head(i) Then
                f.g1.Row = 1
                Iob Mid(Format(f.g1.Text, tip(i)), 1, br(i) + 1) & " ", br(i) + 2, 1'right justify
                GoTo totali1
            End If
        Next x
        Iob "", br(i) + 2, 10
totali1:
    Next i
End If

On Error GoTo tg_print_err1
tg1_print1_1:
For i = 0 To j - 1
    If suma(i) <> -1 Then
        Iob Mid(Format(Format(suma(i), "Standard"), tip(i)), 1, br(i) + 1) & " ", br(i) + 2, 1'right justify
        'st = Format(suma(i), "Standard")
        'printer.CurrentX = xpoz(i + 1) - printer.TextWidth(st)
        'printer.Print st;
    Else
        Iob "", br(i) + 2, 10
    End If
    
Next i
    
prnclose

'---------------------
'DoEvents
Unload tgpr
mp 0
h ""
SPOOL "tgprint"

Exit Sub

tg_print_err1:

'greska2
On Error Resume Next
'MsgBox "Greška pri ispisu!"
Unload tgpr
mp 0
h ""
Exit Sub
End Sub

Sub tg_print2 (f As Form, t As TrueGrid, d As Data, m%)
If draftispis Then tg_print1 f, t, d, m: Exit Sub
'Dim mm%
'mm = m
mp 11
'On Error GoTo tg_print_err2
Dim i%

Load tgpr
tgpr.Combo1.Clear

punitgpr f, t, d'napuni formu za podešavanje s podacima iz truegrid-a
If imazapis(m, "") Then punitgprizbaze m, ""

'prikaz forme za podešavanje ispisa
CenterForm tgpr
tgpr.Show 1
If Len(tgpr.Tag) > 0 Then 'GoTo tg_print_err2 'izlazak-odabrano odustani
    printer.EndDoc
    Unload tgpr
    mp 0
    Exit Sub
End If
    
Dim fn$, fs$
fn = tgpr.panf1.Caption'font name
fs = tgpr.panf2.Caption'font size

'----------------------------
If tgpr.snimi.Value And Len(tgpr.Combo1.Text) > 0 Then upisizapis m, Left(tgpr.Combo1.Text, 30), t.Columns
mp 11
On Error GoTo tg_print_err2

'priprema za ispis
Dim polja() As String, j%, xpoz&(), p&, head()  As String, tgpoz&(), suma#()
'Dim p1&, p2&
j = 0
p = 0
For i = 1 To t.Columns
If tgpr.check(i) Then
   ReDim Preserve polja(j)
   ReDim Preserve head(j)
   ReDim Preserve xpoz(j + 1)
   ReDim Preserve tgpoz(j + 1)
   ReDim Preserve suma(j)
   polja(j) = t.ColumnField(i)
   head(j) = t.ColumnName(i)
   tgpoz(j) = i
   If tgpr.chksum(i) Then suma(j) = 0 Else suma(j) = -1
   'odreðivanje velièine kolona
   'p1 = tgpr.mask(i) * printer.TextWidth("0")'po postavci
   'p2 = printer.TextWidth(head(j)) + printer.TextWidth("0")'da stane heder
   'p = p + IIf(p1 > p2, p1, p2)
   p = p + DajSirinuKolone((tgpr.mask(i)), head(j))
   xpoz(j + 1) = p
   j = j + 1
End If
Next i
If xpoz(j) > printer.ScaleWidth Then
    screen.MousePointer = 0
    If MsgBox("Ispis preko rubova papira! Želite li da nastavim s ispisom?", 36, "Upozorenje") = 7 Then
        Unload tgpr
        Exit Sub
    End If
    screen.MousePointer = 11
End If
'ispis zaglavlja
printer.FontName = fn
printer.FontSize = CDbl(fs) * 1.5
printer.Print tgpr.Text1.Text
printer.Print tgpr.Text2.Text': printer.Print
printer.FontSize = fs
If tgpr.di Then
    printer.Print "Datum ispisa: " & Format(Now, "dd.mm.yyyy"): printer.Print
End If


Dim s As snapshot, strana%
Set s = d.Recordset.CreateSnapshot()'kreiranje snapshota
Dim k%, y%
Dim st As String
strana = 1


ispis_naziva_polja2:
printer.CurrentY = printer.CurrentY + 700

''printer.FontUnderline = True
'''ispis naziva polja
''For i = 0 To j - 1
''   st = head(i)
''   printer.CurrentX = xpoz(i)
''   If s.Fields(polja(i)).Type = 7 Then
''        printer.CurrentX = xpoz(i + 1) - printer.TextWidth(st)
''   End If
''   If i > 0 Then
''        If s.Fields(polja(i - 1)).Type = 7 And s.Fields(polja(i)).Type <> 7 Then printer.CurrentX = xpoz(i) + printer.TextWidth("0")
''   End If
''   printer.Print st;
''Next i
printer.Print
printer.FontUnderline = False
'ispis polja iz baze

Do Until s.EOF
DoEvents
For i = 0 To j - 1
   printer.CurrentX = xpoz(i)
   st = ""
        If IsNull(s.Fields(polja(i)).Value) Then
         st = ""
        Else
         If s.Fields(polja(i)).Type = 7 Then st = Format(s.Fields(polja(i)).Value, "Standard") Else st = s.Fields(polja(i)).Value
         'za logièka polja DA-NE
         If s.Fields(polja(i)).Type = 1 Then
            If s.Fields(polja(i)).Value = 0 Then st = "Ne" Else st = "Da"
         End If

        End If
   'rezanje predugaèkog teksta
   Do Until printer.TextWidth(st) < (xpoz(i + 1) - xpoz(i))
      st = Left(st, Len(st) - 1)
   Loop
   'On Error Resume Next
   If s.Fields(polja(i)).Type = 7 Then
    printer.CurrentX = xpoz(i + 1) - printer.TextWidth(st)
   End If
   If i > 0 Then
    If s.Fields(polja(i - 1)).Type = 7 And s.Fields(polja(i)).Type <> 7 Then printer.CurrentX = xpoz(i) + printer.TextWidth("0")
   End If
   printer.Print st;
   'sumiraj
    If suma(i) <> -1 Then
        suma(i) = suma(i) + db(s.Fields(polja(i)))
    End If
Next i
printer.Print
s.MoveNext
If (printer.CurrentY > printer.ScaleHeight - 5 * printer.TextHeight("stranica " & strana & "   ")) Then
    printer.CurrentX = printer.ScaleWidth - printer.TextWidth("stranica " & strana & "   ")
    printer.CurrentY = printer.ScaleHeight - printer.TextHeight("stranica " & strana & "   ")
    'printer.Print "stranica " & strana;
    strana = strana + 1
   ' If s.EOF Then
   '     h ""
   ' Else
        h "Stranica: " & strana
        printer.NewPage
        GoTo ispis_naziva_polja2
   ' End If
End If
Loop

Dim x%
If tg_suma_g1 = 0 Then GoTo tg1_print_12'preskoèi sumu iz g1
On Error GoTo tg1_print_12
If f.g1.Visible Then
On Error Resume Next
printer.FontBold = True
    For i = 0 To j - 1
        For x = 0 To f.g1.Cols - 1
            f.g1.Row = 0
            f.g1.Col = x
            If f.g1.Text = head(i) Then
                f.g1.Row = 1
                st = f.g1.Text
                printer.CurrentX = xpoz(i + 1) - printer.TextWidth(st)
                printer.Print st;
                GoTo totali2
            End If
        Next x
totali2:
    Next i
End If

tg1_print_12:
printer.FontUnderline = True
For i = 0 To j - 1
    If suma(i) <> -1 Then
        st = Format(suma(i), "Standard")
        printer.CurrentX = xpoz(i + 1) - printer.TextWidth(st)
        printer.Print st;
    End If
Next i
printer.FontUnderline = False
If strana > 1 Then
    printer.CurrentX = printer.ScaleWidth - printer.TextWidth("stranica " & strana & "   ")
    printer.CurrentY = printer.ScaleHeight - printer.TextHeight("stranica " & strana & "   ")
    printer.Print "stranica " & strana;
End If
On Error GoTo tg_print_err2
DoEvents
printer.EndDoc
Unload tgpr
mp 0
h ""
Exit Sub
tg_print_err2:
'MsgBox Error, 16, "Greška"
On Error Resume Next
printer.EndDoc
Unload tgpr
h ""
mp 0
Exit Sub
End Sub

Sub upisizapis (m%, ByVal id$, ByVal br%)
On Error GoTo uzerr
Dim i%
BeginTrans
Dim polj$, sir$, suma$

If Len(Dir$(TgPrintBaz)) > 0 Then
    Dim s As dynaset, b As database
    Set b = OpenDatabase(TgPrintBaz)
    Set s = b.CreateDynaset("select *  from tg_print where m=" & m & " and id='" & id & "'")

For i = 1 To br
    If tgpr.check(i) Then polj = polj & "1" Else polj = polj & "0"
    If tgpr.chksum(i) Then suma = suma & "1" Else suma = suma & "0"
    sir = sir & Format(tgpr.mask(i).Text, "00")
Next i

    If s.RecordCount > 0 Then s.Edit  Else s.AddNew
        s!id = id
        s!m = m
        s!zaglavlje = tgpr.Text1.Text
        s!zaglavlje1 = tgpr.Text2.Text
        s!polja = polj
        s!sirina = sir
        s!suma = suma
        If Not draftispis Then
            s!fn = tgpr.panf1.Caption
            s!fs = tgpr.panf2.Caption
            Load frmprnodabir  'upis orjentacije papira
            s!PORT = frmprnodabir.optOrj(0)
            Unload frmprnodabir
        End If
        s.Update
End If
CommitTrans
Exit Sub
b.Close
s.Close
uzerr:
   greska
   Rollback
   Exit Sub
End Sub

