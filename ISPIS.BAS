Option Explicit
Global PrinterPort%
Global PrinterFileName$
Global BezKodova%
Global NoviKodovi%
Global DraftIspis%
Global HRSlova%

Sub bold (id%)
   If BezKodova Then Exit Sub
   On Error Resume Next
   Select Case id
      Case 0
         Print #1, Chr$(27) & "H" & Chr$(27) & "F";
      Case 1
         Print #1, Chr$(27) & "G" & Chr$(27) & "F";
      Case 2
         Print #1, Chr$(27) & "H" & Chr$(27) & "E";
      Case 3
         Print #1, Chr$(27) & "G" & Chr$(27) & "E";
         
   End Select
End Sub

Sub cpi (v%)
   If BezKodova Then Exit Sub
   On Error Resume Next
   If NoviKodovi Then
      'ESC ! (n)
      'n=0   10 cpi
      'n=1   12 cpi
      'n=4   kondensed
      'n=8   osjenèano
      'n=16  naglašeno
      'n=32  dvostruka širina
      'n=64  proporcionalno
      Select Case v
         Case 5
            Print #1, Chr$(27) & "!" & Chr$(32);
         Case 6
            Print #1, Chr$(27) & "!" & Chr$(36);
         
         Case 7
            Print #1, Chr$(27) & "!" & Chr$(32);
         Case 8
            Print #1, Chr$(27) & "!" & Chr$(36);
         
         Case 10
            Print #1, Chr$(27) & "!" & Chr$(0);
         Case 12
            Print #1, Chr$(27) & "!" & Chr$(1);
         
         Case 17
            Print #1, Chr$(27) & "!" & Chr$(4);
         Case 20
            Print #1, Chr$(27) & "!" & Chr$(5);
      End Select
   Else
      Select Case v
         Case 5
            Print #1, Chr$(27) & "P" & Chr$(27) & "W" & Chr$(1) & Chr$(18);
         Case 6
            Print #1, Chr$(27) & "M" & Chr$(27) & "W" & Chr$(1) & Chr$(18);
         
         Case 7
            Print #1, Chr$(27) & "P" & Chr$(27) & "W" & Chr$(1) & Chr$(15);
         Case 8
            Print #1, Chr$(27) & "M" & Chr$(27) & "W" & Chr$(1) & Chr$(15);
         
         Case 10
            Print #1, Chr$(27) & "P" & Chr$(27) & "W" & Chr$(0) & Chr$(18);
         Case 12
            Print #1, Chr$(27) & "M" & Chr$(27) & "W" & Chr$(0) & Chr$(18);
         
         Case 17
            Print #1, Chr$(27) & "P" & Chr$(27) & "W" & Chr$(0) & Chr$(15);
         Case 20
            Print #1, Chr$(27) & "M" & Chr$(27) & "W" & Chr$(0) & Chr$(15);
      End Select
   End If
End Sub

Function hr (naziv$, tip%)
   If BezKodova Then hr = naziv: Exit Function
   On Error Resume Next
   Dim naziv2$, i%
   If tip = -1 Then hr = naziv: Exit Function
   Select Case tip
      Case 0 'È->C
         For i = 1 To Len(naziv)
            Select Case Mid(naziv, i, 1)
               Case "È": naziv2 = naziv2 & "C"
               Case "è": naziv2 = naziv2 & "c"
               Case "Æ": naziv2 = naziv2 & "C"
               Case "æ": naziv2 = naziv2 & "c"
               Case "Ð": naziv2 = naziv2 & "D"
               Case "ð": naziv2 = naziv2 & "d"
               Case "Š": naziv2 = naziv2 & "S"
               Case "š": naziv2 = naziv2 & "s"
               Case "Ž": naziv2 = naziv2 & "Z"
               Case "ž": naziv2 = naziv2 & "z"
               Case Else: naziv2 = naziv2 & Mid(naziv, i, 1)
            End Select
         Next
      Case 1 'croscii
         For i = 1 To Len(naziv)
            Select Case Mid(naziv, i, 1)
               Case "È": naziv2 = naziv2 & "^"
               Case "è": naziv2 = naziv2 & "~"
               Case "Æ": naziv2 = naziv2 & "]"
               Case "æ": naziv2 = naziv2 & "}"
               Case "Ð": naziv2 = naziv2 & "\"
               Case "ð": naziv2 = naziv2 & "|"
               Case "Š": naziv2 = naziv2 & "["
               Case "š": naziv2 = naziv2 & "{"
               Case "Ž": naziv2 = naziv2 & "@"
               Case "ž": naziv2 = naziv2 & "`"
               
               Case Else: naziv2 = naziv2 & Mid(naziv, i, 1)
            End Select
         Next
      Case 2'code page 852 (East Europe)
         For i = 1 To Len(naziv)
            Select Case Mid(naziv, i, 1)
               Case "È": naziv2 = naziv2 & Chr(&HAC)
               Case "è": naziv2 = naziv2 & Chr(&H9F)
               Case "Æ": naziv2 = naziv2 & Chr(&H8F)
               Case "æ": naziv2 = naziv2 & Chr(&H86)
               Case "Ð": naziv2 = naziv2 & Chr(&HD1)
               Case "ð": naziv2 = naziv2 & Chr(&HD0)
               Case "Š": naziv2 = naziv2 & Chr(&HE6)
               Case "š": naziv2 = naziv2 & Chr(&HE7)
               Case "Ž": naziv2 = naziv2 & Chr(&HA6)
               Case "ž": naziv2 = naziv2 & Chr(&HA7)
          
               Case Else: naziv2 = naziv2 & Mid(naziv, i, 1)
            End Select
         Next
   End Select
   hr = naziv2
End Function

Sub Iob (naziv$, SirKol%, just%)
   Dim str1$, j%
   On Error Resume Next
   naziv = hr(naziv, HRSlova)
   Select Case just
      Case 0 'left
         Print #1, " "; naziv; Spc(SirKol - Len(naziv) - 1);
      Case 10 'left
         Print #1, naziv; Spc(SirKol - Len(naziv));
      Case 1 'right
         Print #1, Spc(SirKol - Len(naziv)); naziv;
      Case 2 'center
         Print #1, Spc(Int(SirKol - Len(naziv)) / 2); naziv; Spc(Int(SirKol - Len(naziv)) / 2);
      Case 5 'repeat
         For j = 1 To SirKol: str1 = str1 & naziv: Next
         Print #1, str1;
      Case 100'left
        Print #1, Spc(SirKol); naziv;
   End Select

End Sub

Sub lq (v%)
   If BezKodova Then Exit Sub
   'If font Then Exit Sub
   'If Not LQdef Then v = 0'default allways off
   On Error Resume Next
   Print #1, Chr$(27) & "x" & Chr$(v);
End Sub

Sub NLine ()
   On Error Resume Next
   Print #1, ""
End Sub

Sub NovaStrana (Strana%)
   cpi 12
   If Strana = 1 Then Iob "stranica " & Strana, 95, 1
   Strana = Strana + 1
   h "stranica " & Strana
   Print #1, Chr(12)
   If PrinterPort < 100 Then
      DoEvents
      
      'Do
      '   DoEvents
      'Loop Until MojMsgBox("Ispis sljedeæe stranice (stranice broj " & Strana & "?", "Dalje ?")
   End If
   Iob "stranica " & Strana, 95, 1
   NLine

End Sub

Sub otvoriprn2 (pl%)
   On Error Resume Next
   mode = -1
   Close
   
   If PrinterPort = 0 Then
      If pl = 1000 Then Open "c:\prnt.txt" For Output As 1: pl = 72 Else Open "prn" For Output As 1
   ElseIf PrinterPort <= 4 Then
      Open "LPT" & Trim(Str$(PrinterPort)) For Output As 1
   Else
      Open PrinterFileName For Output As 1
   End If
   
   If BezKodova Then Exit Sub
   On Error GoTo gre22
   
   Print #1, Chr(27) & "@"; 'reset
   Print #1, Chr(27) & "2"; ' 6 lines/inch
   Print #1, Chr(27); "x"; Chr$(0);  'draft
   Print #1, Chr(27) & "C" & Chr(pl);      'page length (lines)
   cpi 10
   
   'If FONT Then Print #1, Chr(27) & "k" & Chr(FONTtype); ' font
   'Print #1, Chr(27) & "l"; Chr$(ml);   'lijeva margina  na pocetku retka!!!
   'Print #1, Chr(27) & "N" & Chr(mb);     'donja margina
   
   Exit Sub

gre22:
   If MsgBox("Štampaè nije spreman za rad!. Ukljuèite ga i pritisnite OK!", 1, "Greška") <> 1 Then mode = 0: Exit Sub
   Resume

End Sub

Sub prnclose ()
   On Error Resume Next
   If BezKodova Then
      Close
   Else
      Print #1, Chr(12);
      Close
   End If
End Sub

Sub SPOOL (opis$)
   Dim fn$, job$, prn$, port$
   Dim dummy%, prog$
DoEvents
   
   If PrinterPort > 4 Then
      'If BezKodova Then dummy = False Else dummy = MojMsgBox("Ispis na pisaè", "ISPIS")
      dummy = Not BezKodova
If dummy Then
    
    fn = "c:\i" & Left(Trim(Str(Int(Rnd * 10000000))) & "111111111", 7) & ".isp"
    FileCopy PrinterFileName, fn
    DoEvents
    'fn = PrinterFileName'InputBox("Filename", , "c:\tmp.txt")
    job = opis'InputBox("Job", , "AAA")
    'mode = 1'samo odabir printera
    'frmPrnOdabir.Show 1
    
    DoEvents
    'device=HP LaserJet IIP,HPPCL,LPT3:
    Dim default
    default = WinIniDaj("windows", "device", ""):
    If default <> "" Then
       prn = Left(default, InStr(default, ",") - 1)
       port = Mid(default, 1 + InStr(Len(prn) + 2, default, ",")): If Right(port, 1) = ":" Then port = Left(port, Len(port) - 1)
    Else
       prn = InputBox("Printer", , "Epson FX-85")
       port = InputBox("Port", , "LPT1")
    End If
    
    
    dummy = SpoolFile(prn, port, job, fn)
    'MsgBox "spool result : " & Str(SpoolFile(prn, port, job, fn))
Else
    Dim x
    DoEvents
    fn = "c:\i" & Left(Trim(Str(Int(Rnd * 10000000))) & "111111111", 7) & ".pre"
    FileCopy PrinterFileName, fn  '"c:\tmp.txt", "c:\pregled.txt"
    On Error Resume Next
    
    'x = Shell(IIf(FileLen("c:\pregled.txt") < 65000, "Notepad.exe", "SPAD.exe") & " c:\pregled.txt", 1)
    prog = IIf(FileLen(fn) < 32000, "pregled1.exe", "pregled2.exe")'"pregled1.exe"->delphi, "pregled2.exe"->vb za velike dat.(sporo!!!)
    Err = 0
    x = Shell(prog & " " & fn, 1)' " c:\pregled.txt"
    
    'pokušaj sa "Notepad.exe"
    If Err Then
      If FileLen(fn) < 65000 Then
        x = Shell("Notepad.exe" & " " & fn, 1)'" c:\pregled.txt"
      Else
        x = Shell("C:\PROGRA~1\ACCESS~1\WORDPAD.EXE" & " " & fn, 1)
      End If

    End If

    
    DoEvents
    Exit Sub
    'FRMMDI.mnuNOT = True
      End If
   End If
End Sub

Function WinIniDaj (s1$, s2$, s3$)
   Dim x%, temp$
   temp = Space(255)
   'X = GetProfileString(s1, s2, s3, temp, Len(temp), IniFile)
   x = GetProfileString(s1, s2, s3, temp, Len(temp))
   temp = Trim(temp)
   WinIniDaj = Left(temp, Len(temp) - 1)
End Function

