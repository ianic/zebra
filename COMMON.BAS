Option Explicit

Sub CenterForm (x As Form)
On Error Resume Next
    screen.MousePointer = 11
    x.Top = (screen.Height * .95) / 2 - x.Height / 2
    x.Left = screen.Width / 2 - x.Width / 2
    screen.MousePointer = 0
End Sub

Sub CommandMove (c As SSCommand, t$, x!, y!)
    Dim Sir%, Vis%
    Dim z%
    Sir = c.Width
    Vis = c.Height
    If capturestatus = 0 Then
        capturehandle = c.hWnd
        z = SetCapture(capturehandle)
        frm_MDI.SL.Caption = t
        capturestatus = 1
    Else
        If Not (x < 0 Or x >= Sir Or y < 0 Or y >= Vis) Then
            frm_MDI.SL.Caption = t
            Exit Sub
        End If
        ReleaseCapture
        frm_MDI.SL.Caption = ""
        capturestatus = 0
    End If
End Sub

Function DajPunoIme (f As Field)
If FiltPunoIme = 1 Then 'za potrebe složenijih querija sum i sl
    DajPunoIme = f.Name
    Exit Function
End If
   Dim i%
   Dim tmp$
   i = InStr(f.Name, ".")
   If i Then
      tmp = "[" & Left(f.Name, i - 1) & "].[" & Mid(f.Name, i + 1) & "]"
   Else
      If f.SourceTable <> "" Then tmp = "[" & f.SourceTable & "].[" & f.Name & "]" Else tmp = "[" & f.Name & "]"
   End If
   'tmp = IIf(InStr(f.Name, "."), "", "[" & f.SourceTable & "].")
   'tmp = tmp & "[" & f.Name & "]"
   DajPunoIme = tmp
End Function

Sub datkp (M As MaskEdBox, key As Integer)
Select Case key
    Case 68, 100:
            M.Text = Format(Now, "dd.mm.yyyy")
    Case 80, 112:
            M.Text = "01.01." + Format(Now, "yyyy")
    Case 90, 122:
            M.Text = "31.12." + Format(Now, "yyyy")
    Case 43:
         If IsDate(M.Text) Then M.Text = Format(CVDate(M.Text) + 1, "dd.mm.yyyy") Else M.Text = "__.__.____"
    Case 45:
            If IsDate(M.Text) Then M.Text = Format(CVDate(M.Text) - 1, "dd.mm.yyyy") Else M.Text = "__.__.____"
    Case 42:
            If IsDate(M.Text) Then M.Text = Format(CVDate(M.Text) + 7, "dd.mm.yyyy") Else M.Text = "__.__.____"
    Case 47:
            If IsDate(M.Text) Then M.Text = Format(CVDate(M.Text) - 7, "dd.mm.yyyy") Else M.Text = "__.__.____"
    Case 48 To 57
            M.SelLength = 1
End Select
 
End Sub

Sub datlf (M As MaskEdBox)
If Not IsDate(M.Text) Then M.Text = "__.__.____"
End Sub

Function dblkp (i As Integer)
'double textbox polje keypress procedura
If i = 44 Or i = 46 Then
   dblkp = Asc(Format("0", "."))
Else
   dblkp = i
End If
End Function

Sub dbllf (t As TextBox)
If IsNumeric(t.Text) Then t.Text = Format(t.Text, "#,##0.00###") Else t.Text = ""
End Sub

Function df (d)
If IsDate(d) Then
    df = Format(d, "dd.mm.yyyy")
Else
    df = "__.__.____"
End If
End Function

Function dVrijeme (d)

If IsDate(d) Then
    dVrijeme = Format(d, "hh:mm")
Else
    dVrijeme = "__:__"
End If
End Function

Sub h (s$)
frm_MDI.SL = s
End Sub

Function inidaj (s1$, s2$, s3$)
   Dim x%, temp$
   temp = Space(255)
   x = GetPrivateProfileString(s1, s2, s3, temp, Len(temp), Inifile)
   temp = Trim(temp)
   If Asc(Right(temp, 1)) = 0 Then temp = Left(temp, Len(temp) - 1)
   'If temp = s3 Then inipisi s1, s2, s3
   inidaj = temp
End Function

Sub inipisi (s1$, s2$, ByVal s3$)
   Dim x%
   x = OSWritePrivateProfileString(s1$, s2$, s3$, Inifile)
End Sub

Sub intlf (t As TextBox)
'integer textbox polje
If IsNumeric(t.Text) Then t.Text = Format(t.Text, "##0") Else t.Text = ""
End Sub

Sub mp (i%)
screen.MousePointer = i
DoEvents
End Sub

Function NapraviDatum2 (Dat As String)
  Dim LokDat, tmp$
  Dim Kako As String

  Kako = Left$(Dat, 1)
  If Kako = "<" Or Kako = "=" Or Kako = ">" Then
    Dat = Right$(Dat, Len(Dat) - 1)
  Else
    Kako = "="
  End If
  LokDat = DateValue(Dat)
  tmp = "#" & Format$(LokDat, "mm") & "/"
  tmp = tmp & Format$(LokDat, "dd") & "/"
  tmp = tmp & Format$(LokDat, "yyyy") & "#"
  NapraviDatum2 = tmp
End Function

Sub showw (f As Form)
h "Otvaram formu"
f.Show
If f.WindowState = 1 Then f.WindowState = 0
h ""
End Sub

Function StrToDbl (s$)
   Dim tmp$, st$, sd$, znak$, i%
   If IsNumeric(s) Then
      tmp = s' Format(s, "0.00###########")
   Else
      'izbaci separatore 1000
      st = Mid(Format(1000, "#,##0.00"), 2, 1)'separator 1000
      sd = Format(0, ".")'separator - decimalni
      tmp = ""
      For i = 1 To Len(s)
         znak = Mid(s, i, 1)
         If znak >= "0" And znak <= "9" Then
            tmp = tmp & znak
         ElseIf znak = sd Then
            tmp = tmp & znak
            sd = ""
         End If
         'tmp = tmp & IIf(Mid(tmp, i, 1) = st, "", Mid(tmp, i, 1))
      Next
   End If
   StrToDbl = tmp
End Function

