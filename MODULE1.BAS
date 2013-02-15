Option Explicit
Global b As database, str1$, putbaza$, mode%, putbaza1$, str2$
Global tgBC As Long'true grid backcolor
Global backup$

Sub font (fs%, fb%, fi%)
printer.FontSize = fs
printer.FontBold = fb
printer.FontItalic = fi
End Sub

Sub isp (s$, x%, y%)

printer.CurrentX = x
printer.CurrentY = y
printer.Print s;
End Sub

Sub ispisracuna (br As Long, rn%, kopi%)
If IsGotovina() Then
    ispisracunaGotovina br, rn, kopi
    Exit Sub
End If

mp 11
h "Ispis raèuna " & br
Dim x%, y%, pdv%
'vaðenje podataka o pomaku iz ini file-a
x = CInt(inidaj("Ispis", "XPomak", "-130"))
y = CInt(inidaj("Ispis", "YPomak", "300"))
'otvaranje baza faktura i partneri
Dim t As table
Set t = b.OpenTable("faktura")
t.Index = "PrimaryKey"
t.Seek "=", br
If t.NoMatch Then Exit Sub
    Dim t1 As snapshot
    Set t1 = b.CreateSnapshot("select * from partneri where sifra_kupca='" & t!sifra_kupca & "'")
    't1.Index = "PrimaryKey"
    't1.Seek "=", t!sifra_kupca
    If t1.RecordCount = 0 Then Exit Sub
        'postavi orijentaciju papira i broj kopija
        Load frmPrnOdabir
        mp 11
        frmPrnOdabir.optOrj(0) = True
        frmPrnOdabir.cmdPrihvati = True
        frmPrnOdabir.edtCopies = kopi
        Unload frmPrnOdabir
        printer.EndDoc
        'Load postavke
        'postavke.cmd.Copies = kopi
        'Unload postavke
        mp 11
        'izaberi font
        printer.FontName = "Times New Roman CE"
        font 12, 0, 0
        On Error Resume Next
        Dim s1$, s2$, l1&
        Dim s As snapshot
        Dim i As Integer
        If IsPonuda() Then

Else

'podaci o kupcu
ispisxy "Kupac:", 1551 + x, 2000 + y
ispisxy ds(t1!naziv_1), 1551 + x, 2700 + y
ispisxy ds(t1!naziv_2), 1551 + x, 3000 + y
ispisxy ds(t1!adresa), 1551 + x, 3300 + y
ispisxy ds(t1!mjesto), 1551 + x, 3700 + y

'izdavaštvo
font 12, -1, -1
ispisxy TextIznadBrojaRacuna(), 8000 + x, 3650 + y

font 20, -1, 0
'broj raèuna  & "/1"   & "/2"
Dim broj As String
broj = (t!br_racun) & TextporedBrojaRacuna()
ispisxy broj, 9000 - printer.TextWidth(broj) / 2 + x, 3050 + y


'ispis br. narudžbe
font 11, 0, 0
Dim nar As String
nar = Left(t!nap, 3)

If nar = "nar" Then
   ' ispisxy "Narudžbenica br.", 1551 + x, 5500 + y
   nar = Left(t!nap, 15)
   i = Len(nar)
   nar = Right(nar, i - 3)
     ispisxy nar, 3100 + x, 5200 + y
End If


'porezni broj
ispisxy ds(t1!por_br), 2900 + x, 5700 + y
'ispisxy "JMBG:1907942302826", 8300 + x, 5300 + y




ispisxy Format(t!datum, "dd.mm.yyyy. hh:mm") & "", 7600 + x, 4200 + y
ispisxy (t!datum) & ".", 8150 + x, 4700 + y
ispisxy (t!valuta) & ".", 8150 + x, 5200 + y


'ispsi robe
Set s = b.CreateSnapshot("select * from prodroba where br_racun=" & br)

Dim uk(), p() As Double
Dim strana%
ReDim uk(1 To 4)
ReDim p(1 To 4)

i = 7100
strana = 1

Do Until s.EOF

If s!kolicina <> 0 Then
   font 11, 0, 0
   ispisxy (Left(s!nazivrobe, 35)), 1400 + x, y + i
   pdv = s!pdvp
   ispisxy (pdv) & "%", 5450 + x, y + i
   ispisxy (s!jed_mjere), 6150 + x, y + i
   s1 = Format(s!kolicina, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 7700 + x - l1, y + i
   s1 = Format(s!cbp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 8800 + x - l1, y + i
   s1 = Format(s!rabp): l1 = printer.TextWidth(s1)
   ispisxy s1, 9700 + x - l1, y + i
   s1 = Format(s!vbp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 11250 + x - l1, y + i
Else ispisxy (s!nazivrobe), 1400 + x, y + i
End If

   If i > 15000 Then
       font 9, 0, 0
       ispisxy "stranica " & strana, 10300, 15700
       printer.NewPage
       i = 3000
   font 12, 0, 0
   ispisxy "___________________________________________________________________________________", 1400 + x, y + i
   i = i + 300
       strana = strana + 1
   Else i = i + 300
   End If
   

   

    If pdv = 25 Then
        uk(1) = uk(1) + s!vbp - s!vbp / 100 * s!rabp
        End If
    
    If pdv = 23 Then
        uk(2) = uk(2) + s!vbp - s!vbp / 100 * s!rabp
        End If
      
    If pdv = 10 Then
        uk(3) = uk(3) + s!vbp - s!vbp / 100 * s!rabp
       End If
    If pdv = 0 Then
        uk(4) = uk(4) + s!vbp - s!vbp / 100 * s!rabp
    End If
  s.MoveNext
Loop

'If t!rabi > 0 Then
'   font 12, 0, 0
'   ispisxy "- " & Format(t!rabp, "Standard") & "% rabata", 6750 + x, 10850 + y
'   font 14, 0, 0
'   ispisxy Format(t!rabi, "Standard"), 9650 + x + 1700 - printer.TextWidth(Format(t!rabi, "standard")), 10850 + y
'End If


   If i > 13000 Then
       font 9, 0, 0
         ispisxy "stranica " & strana, 10300, 15700
       printer.NewPage
      i = 3000
       strana = strana + 1
   End If




'TOTALI
   font 12, 0, 0
   
   ispisxy "___________________________________________________________________________________", 1400 + x, y + i

   s1 = Format(t!vbp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy "Iznos naknade", 8000, 400 + y + i
   ispisxy s1, 11250 + x - l1, 400 + y + i
   
   s1 = Format(t!rab, "standard"): l1 = printer.TextWidth(s1)
   ispisxy "Iznos rabata", 8000, 800 + y + i
   ispisxy s1, 11250 + x - l1, 800 + y + i
   
   s1 = t!vbp - t!rab
   s1 = Format(s1, "standard"): l1 = printer.TextWidth(s1)
   ispisxy "Iznos netto", 8000, 1200 + y + i
   ispisxy s1, 11250 + x - l1, 1200 + y + i
   
   s1 = Format(t!pdv, "standard"): l1 = printer.TextWidth(s1)
   ispisxy "Iznos PDV-a", 8000, 1600 + y + i
   ispisxy s1, 11250 + x - l1, 1600 + y + i
  
   font 12, -1, 0
   s1 = Format(t!vsp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy "Ukupno", 8000, 2000 + y + i
   ispisxy s1, 11250 + x - l1, 2000 + y + i: printer.FontBold = False

   'Dim t2 As table
   'Set t2 = b.OpenTable("Porezi")
   't2.Index = "PrimaryKey"
   't2.Seek "=", t!poro
   'font 12, 0, 0
   'ispisxy (t2!opis1), 3000 + x, 12900
   'ispisxy (t2!opis2), 3000 + x, 13200
   'font 12, 0, 0
 'ispisxy  "  " 3000 + x, 13000
 
 ' obraèun poreza po stopama
 font 11, 0, 0
 If uk(1) <> 0 Then
   p(1) = uk(1) * .25
   s1 = Format(uk(1), "standard"): l1 = printer.TextWidth(s1)
   s2 = Format(p(1), "standard"): l1 = printer.TextWidth(s1)
   ispisxy "osnovica 25% = " & s1 & " , PDV 25% = " & s2, 1400 + x, 600 + y + i
 End If

 If uk(2) <> 0 Then
   p(2) = uk(2) * .23
   s1 = Format(uk(2), "standard"): l1 = printer.TextWidth(s1)
   s2 = Format(p(2), "standard"): l1 = printer.TextWidth(s1)
   ispisxy "osnovica 23% = " & s1 & " , PDV 23% = " & s2, 1400 + x, 900 + y + i
 End If

 If uk(3) <> 0 Then
   p(3) = uk(3) * .1
   s1 = Format(uk(3), "standard"): l1 = printer.TextWidth(s1)
   s2 = Format(p(3), "standard"): l1 = printer.TextWidth(s1)
   ispisxy "osnovica 10% = " & s1 & " , PDV 10% = " & s2, 1400 + x, 1200 + y + i
 End If

 If uk(4) <> 0 Then
   s1 = Format(uk(4), "standard"): l1 = printer.TextWidth(s1)
  ispisxy "osnovica 0% = " & s1, 1400 + x, 1500 + y + i
 End If

     
   font 11, 0, 0
   ispisxy TextPozivNaBroj(broj), 1400 + x, 2000 + y + i
   
   font 12, 0, 0
   ispisxy "Potpis:", 7000, 2800 + y + i



 If strana > 1 Then
       font 9, 0, 0
         ispisxy "stranica " & strana, 10300, 15700
 End If


printer.EndDoc

End If




printer.EndDoc
'vrati broj kopija na 1
'postavke.cmd.Copies = 1
'Unload postavke
Load frmPrnOdabir
mp 11
'frmPrnOdabir.optOrj(0) = True
'frmPrnOdabir.cmdPrihvati = True
frmPrnOdabir.edtCopies = 1
Unload frmPrnOdabir

   printer.EndDoc
h ""
mp 0
End Sub

Sub ispisracunaGotovina (br As Long, rn%, kopi%)
                                                                    mp 11
h "Ispis raèuna " & br
Dim x%, y%
'vaðenje podataka o pomaku iz ini file-a
x = CInt(inidaj("Ispis", "XPomak", "-130"))
y = CInt(inidaj("Ispis", "YPomak", "300"))
'otvaranje baza faktura i partneri
Dim t As table
Set t = b.OpenTable("faktura")
t.Index = "PrimaryKey"
t.Seek "=", br
If t.NoMatch Then Exit Sub
Dim t1 As snapshot
Set t1 = b.CreateSnapshot("select * from partneri where sifra_kupca='" & t!sifra_kupca & "'")
't1.Index = "PrimaryKey"
't1.Seek "=", t!sifra_kupca
If t1.RecordCount = 0 Then Exit Sub
'postavi orijentaciju papira i broj kopija
Load frmPrnOdabir
mp 11
frmPrnOdabir.optOrj(0) = True
frmPrnOdabir.cmdPrihvati = True
frmPrnOdabir.edtCopies = kopi
Unload frmPrnOdabir
printer.EndDoc
'Load postavke
'postavke.cmd.Copies = kopi
'Unload postavke
mp 11
'izaberi font
printer.FontName = "Times New Roman CE"
font 12, 0, 0
On Error Resume Next
'podaci o kupcu
ispisxy ds(t1!naziv_1), 2100 + x, 50 + y
ispisxy ds(t1!naziv_2), 2100 + x, 350 + y
ispisxy ds(t1!adresa), 2100 + x, 780 + y
ispisxy ds(t1!mjesto), 2100 + x, 1150 + y
'porezni broj
  ispisxy ds(t1!por_br), 2100 + x, 2100 + y
   'ispisxy "JMBG:1907942302826", 8300 + x, 5250 + y  izbaèeno

'radni nalog
'If rn Then
 '  font 14, -1, -1
  ' ispisxy "Radni nalog", 8950 + x, 3100 + y
'End If
font 18, -1, 0

'broj raèuna,datum,valuta
Dim broj As String
broj = (t!br_racun) & TextporedBrojaRacuna()
ispisxy broj, 3600 - printer.TextWidth((t!br_racun)) / 2 + x, 2650 + y
font 12, 0, 0
ispisxy Format(t!datum, "dd.mm.yyyy. hh:mm") & "", 6100 + x, 2750 + y
ispisxy "31.12.2012.", 9650 + x, 2750 + y

Dim s As snapshot, i As Integer
'ispsi robe
Set s = b.CreateSnapshot("select * from prodroba where br_racun=" & br)
Dim s1$, s2$, l1&

Dim pdv As Integer
Dim uk(), p() As Double
ReDim uk(1 To 3)
ReDim p(1 To 3)

Do Until s.EOF
   font 11, 0, 0
   ispisxy (s!nazivrobe), 800 + x, 4000 + y + i
   ispisxy (s!jed_mjere), 4400 + x, 4000 + y + i
   s1 = Format(s!kolicina, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 5700 + x - l1, 4000 + y + i
   s1 = Format(s!cbp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 6900 + x - l1, 4000 + y + i
   s1 = Format(s!vbp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 8100 + x - l1, 4000 + y + i
   pdv = Format(s!pdvp, "standard"):
   ispisxy pdv & "%", 8400 + x, 4000 + y + i
   s1 = Format(s!pdvi, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 9800 + x - l1, 4000 + y + i
   font 12, -1, 0
   s1 = Format(s!vsp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 11100 + x - l1, 4000 + y + i


    If s!pdvp = 25 Then
        uk(1) = uk(1) + s!vbp
        End If
    If s!pdvp = 10 Then
        uk(2) = uk(2) + s!vbp
       End If
    If s!pdvp = 0 Then
        uk(3) = uk(3) + s!vbp
    End If



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
   s1 = Format(t!vbp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 8150 + x - l1, 9100 + y
   s1 = Format(t!pdv, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 9800 + x - l1, 9100 + y
   font 14, -1, 0
   s1 = Format(t!vsp, "standard"): l1 = printer.TextWidth(s1)
   ispisxy s1, 11100 + x - l1, 9100 + y: printer.FontBold = False

 font 10, 0, 0
 If uk(1) <> 0 Then
   p(1) = uk(1) * .25
   s1 = Format(uk(1), "standard"): l1 = printer.TextWidth(s1)
   s2 = Format(p(1), "standard"): l1 = printer.TextWidth(s1)
   ispisxy "osnovica 25% = " & s1 & " , PDV 25% = " & s2, 800 + x, 10000
 End If

 If uk(2) <> 0 Then
   p(2) = uk(2) * .1
   s1 = Format(uk(2), "standard"): l1 = printer.TextWidth(s1)
   s2 = Format(p(2), "standard"): l1 = printer.TextWidth(s1)
   ispisxy "osnovica 10% = " & s1 & " , PDV 10% = " & s2, 800 + x, 10300
 End If

 If uk(3) <> 0 Then
   s1 = Format(uk(3), "standard"): l1 = printer.TextWidth(s1)
  ispisxy "osnovica 0% = " & s1, 800 + x, 10600
 End If




   'Dim t2 As table
   'Set t2 = b.OpenTable("Porezi")
   't2.Index = "PrimaryKey"
   't2.Seek "=", t!poro
   'font 12, 0, 0
   'ispisxy (t2!opis1), 3000 + x, 12900
   'ispisxy (t2!opis2), 3000 + x, 13200
   font 12, 0, 0
 ' ispisxy "RAÈUN PLAÆEN GOTOVINOM", 900 + x, 9500
printer.EndDoc
'vrati broj kopija na 1
'postavke.cmd.Copies = 1
'Unload postavke
Load frmPrnOdabir
mp 11
'frmPrnOdabir.optOrj(0) = True
'frmPrnOdabir.cmdPrihvati = True
frmPrnOdabir.edtCopies = 1
Unload frmPrnOdabir

   printer.EndDoc
h ""
mp 0
End Sub

Sub ispisxy (s$, x%, y%)
printer.CurrentX = x
printer.CurrentY = y
printer.Print s;
End Sub

Sub Main ()
On Error Resume Next
'pobriši datoteke nastale zbog pregleda
Kill "c:\i*.pre"

On Error GoTo mainexit

inifile = app.Path & "\" & app.EXEName & ".ini"

TgPrintBaz = inidaj("Baze", "PutDoTgprint", (app.Path) & "\tg_print.mdb")
backup = inidaj("Baze", "BackupDir", (app.Path) & "\baze")

DraftIspis = 0'CInt(inidaj("Ispis", "Draft", "0"))
BezKodova = CInt(inidaj("Ispis", "BezKodova", "-1"))'ispis na ekran
IspisSve = CInt(inidaj("Ispis", "IspisSve", "-1"))'detaljan ispis raèuna ili ne

PrinterPort = inidaj("Ispis", "PrinterPort", "100")
NoviKodovi = CInt(inidaj("Ispis", "NoviKodovi", "0"))
HRSlova = CInt(inidaj("Ispis", "HRSlova", "1"))'za moj printer 1 za Nenin 2

PrinterFileName = app.Path & "\tmp_isp.txt"
tgBC = CLng(inidaj("Boje", "tablice", Str(12632256)))

Dim s$, p$

p = inidaj("Baze", "PutDoBaze", (app.Path) & "\baze")

s = Dir$(p & "\*.mdb", 0)

putbaza = p & "\" & s
putbaza1 = p
s = Dir$
If Len(s) > 0 Then baze.Show  Else frm_mdi.Show

Exit Sub
mainexit:
   MsgBox Error, 16, "Greška"
   Exit Sub
End Sub

Sub MakniItalic (f As Form)
On Error Resume Next
Dim i
For i = 0 To f.Controls.Count - 1
   f.Controls(i).FontItalic = False
Next i
End Sub

Sub NamjestiTGBC (f As Form)
Dim i%
For i = 0 To f.Controls.Count - 1
If TypeOf f.Controls(i) Is TrueGrid Then
    f.Controls(i).HeadBackColor = tgBC
    f.Controls(i).InactiveBackColor = tgBC
    DoEvents
End If
Next i
On Error Resume Next
'f.ch1.BackColor = f.BackColor
'f.Ch2.BackColor = f.BackColor
End Sub

