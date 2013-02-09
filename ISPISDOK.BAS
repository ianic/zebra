Option Explicit
Global ispissve%

Function IsGotovina ()
    IsGotovina = (inidaj("Ispis", "Gotovina", "") <> "")
End Function

Sub ispisIRG (br As Long)
mp 11
h "Ispis raèuna " & br
Dim x%, y%
'vaðenje podataka o pomaku iz ini file-a
x = CInt(inidaj("Ispis", "XPomak1", "0"))
y = CInt(inidaj("Ispis", "YPomak1", "0"))
'otvaranje baza faktura i partneri
Dim t As table
Set t = b.OpenTable("faktura")
t.Index = "PrimaryKey"
t.Seek "=", br
If t.NoMatch Then Exit Sub
Dim t1 As table
Set t1 = b.OpenTable("partneri")
t1.Index = "PrimaryKey"

t1.Seek "=", 0
If t1.NoMatch Then MsgBox "Podaci o matiènoj firmi nisu upisani (partner sa šifrom 0)!!!", 16, "": Exit Sub

'postavi orijentaciju papira
Load frmPrnOdabir
frmPrnOdabir.optOrj(0) = True
frmPrnOdabir.cmdPrihvati = True
Unload frmPrnOdabir
printer.EndDoc

'izaberi font
printer.FontName = "Arial CE"
font 14, 1, 0
ispisxy ds(t!tip), 10000 + x, 100 + y
font 10, 0, 0
ispisxy ds(t1!naziv_1), 100 + x, 100 + y
ispisxy ds(t1!naziv_2), 100 + x, 350 + y
ispisxy ds(t1!adresa), 100 + x, 600 + y
ispisxy ds(t1!mjesto), 100 + x, 850 + y
ispisxy "Žiro raèun: " & ds(t1!ziro), 100 + x, 1100 + y

printer.FontName = "Times New Roman CE"
font 12, 0, 0
On Error Resume Next
'podaci o kupcu
t1.Seek "=", t!sifra_kupca
If t1.NoMatch Then Exit Sub
ispisxy ds(t1!naziv_1), 500 + x, 1700 + y
ispisxy ds(t1!naziv_2), 500 + x, 2000 + y
ispisxy ds(t1!adresa), 500 + x, 2300 + y
ispisxy ds(t1!mjesto), 500 + x, 2600 + y
'porezni broj
ispisxy "Porezni broj: " & ds(t1!por_br), 1500 + x, 3200 + y
'radni nalog
'If rn Then
'   font 14, -1, -1
'   ispisxy "Radni nalog", 8950 + x, 3100 + y
'End If
font 14, -1, 0
'broj raèuna,datum,valuta

ispisxy "Raèun: " & (t!br_racun), 8700 + x, 2700 + y
font 12, 0, 0
ispisxy "Datum: " & Format(ds(t!datum), "dd.mm.yyyy"), 8000 + x, 2950 + y
ispisxy "Valuta: " & Format(ds(t!valuta), "dd.mm.yyyy"), 8000 + x, 3250 + y
font 12, 0, 0
ispisxy "PDV %", 4000 + x, 4900 + y

Dim s As snapshot, i As Integer
'ispsi robe
'printer.FontName = "Arial CE"
Set s = b.CreateSnapshot("select * from prodroba where br_racun=" & br)
Dim s1$, l1&
i = 5000
Do Until s.EOF
   font 12, 0, 0
   ispisxy ds(s!nazivrobe), 100 + x, y + i
   s1 = Format(s!pdvp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 6000 + x - l1, y + i
   ispisxy ds(s!jed_mjere), 6500 + x, y + i
   s1 = Format(s!kolicina, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 8500 + x - l1, y + i
   s1 = Format(s!cbp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 10000 + x - l1, y + i
   font 14, 0, 0
   s1 = Format(s!vbp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 11500 + x - l1, y + i
   i = i + 300
   s.MoveNext
Loop

'If t!rabi > 0 Then
'   font 12, 0, 0
'   ispisxy "- " & Format(t!rabp, "Standard") & "% rabata", 6750 + x, 10850 + y
'   font 14, 0, 0
'   ispisxy Format(t!rabi, "Standard"), 9650 + x + 1700 - printer.TextWidth(Format(t!rabi, "standard")), 10850 + y
'End If

'TOTALI
   font 14, 0, 0
   i = i + 200
   s1 = Format(t!vbp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 11500 + x - l1, i + y
   s1 = Format(t!rab, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 11500 + x - l1, i + 300 + y
   s1 = Format(t!pdv, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 11500 + x - l1, i + 600 + y
   font 14, -1, 0
   s1 = Format(t!vsp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 11500 + x - l1, i + 900 + y: printer.FontBold = False

   'Dim t2 As table
   'Set t2 = b.OpenTable("Porezi")
   't2.Index = "PrimaryKey"
   't2.Seek "=", t!poro
   'font 12, 0, 0
   'ispisxy (t2!opis1), 3000 + x, 12900
   'ispisxy (t2!opis2), 3000 + x, 13200
   'font 12, 0, 0
   'ispisxy "Kod plaæanja ovog raèuna pozovite se na broj 02 " & br & ".", 3000 + x, 13500
printer.EndDoc
h ""
mp 0

End Sub

Sub ispisur (BrR As Long)
'ispis ulaznog raèuna
Dim t As table, p As table
Set t = b.OpenTable("faktura")
Set p = b.OpenTable("partneri")
t.Index = "PrimaryKey"
p.Index = "PrimaryKey"
t.Seek "=", BrR
If t.NoMatch Then Exit Sub
p.Seek "=", 0
If p.NoMatch Then MsgBox "Podaci o matiènoj firmi nisu upisani (partner sa šifrom 0)!!!", 16, "": Exit Sub
otvoriprn2 72'-----------
If mode = 0 Then Exit Sub
cpi 17
bold 0
'matièna firma
cpi 12
Iob ds(p!naziv_1) & " " & ds(p!naziv_2), 90, 10
cpi 10: bold 1: Iob ds(t!tip), 4, 1: NLine: bold 0: cpi 12
Iob ds(p!adresa), 1, 10: NLine
Iob ds(p!mjesto), 1, 10: NLine
Iob "Tel: " & ds(p!telefon) & "  Fax:" & ds(p!telefax), 1, 10: NLine
Iob "Porezni broj: " & ds(p!por_br), 1, 10: NLine
Iob "Žiro: " & ds(p!ziro), 1, 10: NLine: NLine
'kupac
p.Seek "=", t!sifra_kupca
If Not p.NoMatch Then
cpi 10
Iob ds(p!naziv_1) & " " & ds(p!naziv_2), 20, 100: NLine
Iob ds(p!adresa), 20, 100: NLine
Iob ds(p!mjesto), 20, 100: NLine
Iob "Porezni broj: " & ds(p!por_br), 20, 100: NLine: NLine
End If
bold 1
    cpi 10
    Iob "Raèun broj: " & t!br_racun, 15, 100: NLine
bold 0
cpi 12
Iob "Datum raèuna   : " & dd(t!datum), 15, 100: NLine
Iob "Datum dospijeæa: " & dd(t!valuta), 15, 100: NLine
Iob "=", 96, 5: NLine
Dim s As snapshot
Set s = b.CreateSnapshot("select * from prodroba where br_racun=" & t!br_racun)
cpi 12

'zaglavlje
    Iob "R.B.", 4, 10
    Iob "Naziv robe", 51, 10
    Iob "j.mj.", 5, 10
    Iob "kolièina", 9, 1
    Iob "cijena", 13, 1
    Iob "vrijednost", 13, 1

    If ispissve Then
    NLine
    Iob "rabat", 42, 1
    Iob "", 7, 1
    Iob "PDV", 13, 1
    Iob "", 7, 1
    Iob "cij. s PDV", 13, 1
    bold 1: Iob "vr. s PDV", 13, 1: bold 0
    End If
    NLine
Iob "-", 96, 5: NLine
Dim rb%
rb = 0
While Not s.EOF
    rb = rb + 1
    Iob Trim(Str(rb)) & " ", 4, 1
    Iob ds(s!nazivrobe), 51, 10
    Iob ds(s!jed_mjere), 5, 10
    Iob dbf(s!kolicina), 9, 1
    Iob dbf(s!cbp), 13, 1
    Iob dbf(s!vbp), 13, 1

    If ispissve Then
    NLine
    Iob dbf(s!rabi), 42, 1
    Iob dbf(s!rabp) & "%", 7, 1
    Iob dbf(s!pdvi), 13, 1
    Iob dbf(s!pdvp) & "%", 7, 1
    Iob dbf(s!csp), 13, 1
    bold 1: Iob dbf(s!vsp), 13, 1: bold 0
    End If
    NLine
s.MoveNext
'    If Not s.EOF Then iob "-", 96, 5: nline
Wend
Iob "=", 96, 5: NLine: NLine
cpi 10
bold 1
Iob "", 10, 10: Iob "Vrijednost:", 13, 10: Iob dbf(t!vbp), 13, 1: NLine
If dbf(t!rab) > 0 Then
    Iob "", 10, 10: Iob "Rabat:", 13, 10: Iob dbf(t!rab), 13, 1: NLine
End If
Iob "", 10, 10: Iob "PDV:", 13, 10: Iob dbf(t!pdv), 13, 1: NLine
bold 3
Iob "", 10, 10: Iob "Za naplatu:", 13, 10: Iob dbf(t!vsp), 13, 1: NLine
bold 0
prnclose'-------------
SPOOL "Ispis UR:" & BrR
End Sub

Function IsPonuda ()
  IsPonuda = False'(UCase$(Left$(inidaj("baze", "ZadnjaBaza", ""), 1)) = "P")
End Function

Function TextIznadBrojaRacuna ()
  TextIznadBrojaRacuna = inidaj("Ispis", "TextIznadBrojaRacuna", "")
End Function

Function TextporedBrojaRacuna ()
    TextporedBrojaRacuna = inidaj("Ispis", "TextPoredBrojaRacuna", "")
End Function

Function TextPozivNaBroj (br As String)
    TextPozivNaBroj = ""
    Dim t As String
    t = inidaj("Ispis", "TextPozivNaBroj", "")
    If t <> "" Then
     TextPozivNaBroj = t & br & "."
    End If
End Function

