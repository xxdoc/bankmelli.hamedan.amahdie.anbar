Attribute VB_Name = "Kod"
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Public Hav1$, Kal1$
Public cnn As New Connection, rcd As New Recordset
Private Type Kala
     Name   As String * 50
     Vahed  As String * 30
     Code   As String * 20
     Serial As String * 30
     Mojody As Integer
     Tozih  As String * 50
End Type
Public asd As New FileSystemObject
Public Kala1() As Kala
Public Kalmdb$, Havmdb, Shomdb$
Public i%, Mojaz As Boolean  'baraye por kardan list mojaz ast?

Public Sub LodCombo(ByRef Dcc As Adodc)
On Error GoTo 4
'If Mojaz = True Then Exit Sub
'Mojaz = True
With Dcc
ReDim Kala1(0 To .Recordset.RecordCount - 1) As Kala
For i = 0 To .Recordset.RecordCount - 1
Kala1(i).Name = Null2Str(.Recordset(0).Value)
Kala1(i).Vahed = Null2Str(.Recordset(1).Value)
Kala1(i).Mojody = Null2Str(.Recordset(2).Value)
Kala1(i).Code = Null2Str(.Recordset(3).Value)
Kala1(i).Serial = Null2Str(.Recordset(4).Value)
Kala1(i).Tozih = Null2Str(.Recordset(5).Value)
.Recordset.MoveNext
Next
'--------
.Recordset.MoveFirst
End With
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger("LodCombo() Of Module 'Kod',Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Public Sub Loger(ByVal Log As String)
Dim Ffile As String
On Error GoTo 4 '---------------
Ffile = IIf(Len(App.Path) = 3, App.Path + "log.log", App.Path + "\log.log")
If asd.FileExists(Ffile) = False Then asd.CreateTextFile Ffile
On Error GoTo 5 '---------------
Open Ffile For Append As #2
'FIXIT: Print method has no Visual Basic .NET equivalent and will not be upgraded.         FixIT90210ae-R7593-R67265
Print #2, " - Time:" & Time$ & " - Date:" & Date$ & " - Error:" & Log
Close
Exit Sub
4: MsgBox "«„ﬂ«‰ «ÌÃ«œ ›«Ì· œ— Õ«›ŸÂ ÊÃÊœ ‰œ«—œ .«Õ „«·¬ »—‰«„Â Ì« «“ —ÊÌ ”Ì œÌ «Ã—« „Ì ‘Êœ Ì« œ—«ÌÊ ‰’» »—‰«„Â Å— ‘œÂ «”  ·ÿ›¬ „”Ì— ‰’» »—‰«„Â —« çﬂ ﬂ‰Ìœ!", vbExclamation, "Œÿ« œ— œ” —”Ì »Â Õ«›ŸÂ"
Exit Sub
5: MsgBox "«„ﬂ«‰ ‰Ê‘ ‰ œ— ›«Ì· Œÿ« ÊÃÊœ ‰œ«—œ «Õ „«·¬ ›«Ì· Õ›«Ÿ  ‘œÂ «” .·ÿ›¬ „”Ì— ‰’» »—‰«„Â —« çﬂ ﬂ‰Ìœ", vbExclamation, "Œÿ« œ— ‰Ê‘ ‰ —ÊÌ ›«Ì·"
End Sub
'FIXIT: Declare 'Null2Str' and 'Value' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Private Function Null2Str(Value As Variant) As Variant
If IsNull(Value) = True Then
Null2Str = ""
Else
Null2Str = Value
End If
End Function
'FIXIT: Declare 'Fild2Str' and 'Value' with an early-bound data type                       FixIT90210ae-R1672-R1B8ZE
Public Function Fild2Str(Value As Variant) As Variant
If IsNull(Value) = True Then
Fild2Str = ""
Else
Fild2Str = Value
End If
End Function

Public Function IRDate(Optional ByVal Dat As String) As String
Dim arryDay(1 To 7) As String
Dim arryMon(1 To 12) As String
Dim DD$, MM$, YY$, ii%, Tm$, Td%, RR$
arryDay(1) = "‘‰»Â"
arryDay(2) = "Ìﬂ‘‰»Â"
arryDay(3) = "œÊ‘‰»Â"
arryDay(4) = "”Â ‘‰»Â"
arryDay(5) = "çÂ«—‘‰»Â"
arryDay(6) = "Å‰Ã‘‰»Â"
arryDay(7) = "Ã„⁄Â"
'-----------------
arryMon(1) = "›—Ê—œÌ‰"
arryMon(2) = "«—œÌ»Â‘ "
arryMon(3) = "Œ—œ«œ"
arryMon(4) = " Ì—"
arryMon(5) = "„—œ«œ"
arryMon(6) = "‘Â—ÌÊ—"
arryMon(7) = "„Â—"
arryMon(8) = "¬»«‰"
arryMon(9) = "¬–—"
arryMon(10) = "œÌ"
arryMon(11) = "»Â„‰"
arryMon(12) = "«”›‰œ"
'-----------------
If Dat = "" Then
RR = arryDay(Val(Weekday(Date + 1)))
DD = Day(Date): MM = Month(Date): YY = Year(Date)
Else
RR = Right(Dat, 2) + 1
DD = (Val(Right(Dat, 2)))
MM = Val(Mid(Dat, 6, 2))
YY = Val(Left(Dat, 4))
End If
If DD <= 1 Then
          ii = YY - 622
Else
       If DD <= 19 And MM = 2 Then
          ii = YY - 622
       Else
          ii = YY - 621
       End If
End If
 MM = MM - 2
'if (now.getMonth() <= 1 ) { ii=now.getFullYear()-622 }
'else { if ( now.getDate() <= 21 && now.getMonth() == 2)
'     { ii=now.getFullYear()-622 }
'     else { ii=now.getFullYear()-621 }}
'-----------------------------------------
If MM > 5 And MM < 10 Then Tm = IIf(DD <= 22, arryMon(MM), arryMon(MM + 1))
'if ( now.getMonth() == 6 || now.getMonth() == 7 || now.getMonth() == 8 || now.getMonth() == 9)
'{ if ( now.getDate() <=22 ) { mm=monName[now.getMonth()] } else {mm=monName[now.getMonth()+1]}}
If MM = 4 Or MM = 5 Or MM = 10 Then Tm = IIf(DD <= 21, arryMon(MM), arryMon(MM + 1))
'if ( now.getMonth() == 4 || now.getMonth() == 5 || now.getMonth() == 10 )
'{ if ( now.getDate() <=21 ) { mm=monName[now.getMonth()] } else {mm=monName[now.getMonth()+1] }}
If MM >= 0 And MM < 4 Then Tm = IIf(DD <= 20, arryMon(MM), arryMon(MM + 1))
'if ( now.getMonth() == 0 || now.getMonth() == 2 || now.getMonth() == 3 )
'{ if ( now.getDate() <=20 ) { mm=monName[now.getMonth()] } else { mm=monName[now.getMonth()+1] }}
If MM = 1 Then Tm = IIf(DD <= 19, arryMon(MM), arryMon(MM + 1))
'if ( now.getMonth() == 1) { if ( now.getDate() <=19 ) { mm=monName[now.getMonth()] }
'else { mm=monName[now.getMonth()+1] }}
If MM = 11 Then Tm = IIf(DD <= 21, arryMon(MM), arryMon(MM - 11))
'if ( now.getMonth() == 11) { if ( now.getDate() <=21 ) { mm=monName[now.getMonth()] }
'else { mm=monName[now.getMonth()-11] }}
If MM = 2 Or MM = 3 Or MM = 0 Then Td = IIf(DD <= 20, DD + 10, DD + 10 - 30)
'if ( now.getMonth() == 2 || now.getMonth() == 3 || now.getMonth() == 0 )
'{if ( now.getDate() <=20 ) {mr=now.getDate()+10 } else {  mr=now.getDate()+10-30 }}
If MM = 4 Or MM = 5 Then Td = IIf(DD <= 21, DD + 10, DD + 10 - 31)
'if ( now.getMonth() == 4 || now.getMonth() == 5 )
'{ if ( now.getDate() <=21 ) { mr=now.getDate()+10 } else { mr=now.getDate()+10-31 }}
If MM > 5 And MM < 9 Then Td = IIf(DD <= 21, DD + 9, DD + 9 - 31)
'if ( now.getMonth() == 6 || now.getMonth() == 7 || now.getMonth() == 8 )
'{ if ( now.getDate() <=21 ) { mr=now.getDate()+9 } else { mr=now.getDate()+9-31 }}
If MM = 10 Or MM = 11 Then Td = IIf(DD <= 21, DD + 9, DD + 9 - 30)
'if ( now.getMonth() == 10 || now.getMonth() == 11)
'{ if ( now.getDate() <=21 ) { mr=now.getDate()+9 } else { mr=now.getDate()+9-30 }}
If MM = 9 Then Td = IIf(DD <= 22, DD + 8, DD + 8 - 31)
'if ( now.getMonth() == 9){ if ( now.getDate() <=22 ) { mr=now.getDate()+8 }
'else { mr=now.getDate()+8-30} }
If MM = 1 Then Td = IIf(DD <= 20, DD + 11, DD + 11 - 31)
'if ( now.getMonth() == 1 ){ if ( now.getDate() <=20 ) { mr=now.getDate()+11 }
'else { mr=now.getDate()+11-30  }}
IRDate = RR & "/" & Td & "/" & MM & "/" & ii
'document.write (RR + ", " + mr + " " + MM + " " + ii)
End Function

Public Function IRDate2(Optional ByVal Dat As String) As String
Dim arryDay(1 To 7) As String
Dim arryMon(1 To 12) As String
Dim DD$, MM$, YY$, ii%, Tm$, Td%, RR$
arryDay(1) = ""
arryDay(2) = ""
arryDay(3) = ""
arryDay(4) = ""
arryDay(5) = ""
arryDay(6) = ""
arryDay(7) = ""
'-----------------
arryMon(1) = "›—Ê—œÌ‰"
arryMon(2) = "«—œÌ»Â‘ "
arryMon(3) = "Œ—œ«œ"
arryMon(4) = " Ì—"
arryMon(5) = "„—œ«œ"
arryMon(6) = "‘Â—ÌÊ—"
arryMon(7) = "„Â—"
arryMon(8) = "¬»«‰"
arryMon(9) = "¬–—"
arryMon(10) = "œÌ"
arryMon(11) = "»Â„‰"
arryMon(12) = "«”›‰œ"
'-----------------
If Dat = "" Then
RR = arryDay(Val(Weekday(Date + 1)))
DD = Day(Date): MM = Month(Date): YY = Year(Date)
Else
RR = Right(Dat, 2) + 1
DD = (Val(Right(Dat, 2)))
MM = Val(Mid(Dat, 6, 2))
YY = Val(Left(Dat, 4))
End If
If DD <= 1 Then
          ii = YY - 622
Else
       If DD <= 19 And MM = 2 Then
          ii = YY - 622
       Else
          ii = YY - 621
       End If
End If
 MM = MM - 2
'if (now.getMonth() <= 1 ) { ii=now.getFullYear()-622 }
'else { if ( now.getDate() <= 21 && now.getMonth() == 2)
'     { ii=now.getFullYear()-622 }
'     else { ii=now.getFullYear()-621 }}
'-----------------------------------------
If MM > 5 And MM < 10 Then Tm = IIf(DD <= 22, arryMon(MM), arryMon(MM + 1))
'if ( now.getMonth() == 6 || now.getMonth() == 7 || now.getMonth() == 8 || now.getMonth() == 9)
'{ if ( now.getDate() <=22 ) { mm=monName[now.getMonth()] } else {mm=monName[now.getMonth()+1]}}
If MM = 4 Or MM = 5 Or MM = 10 Then Tm = IIf(DD <= 21, arryMon(MM), arryMon(MM + 1))
'if ( now.getMonth() == 4 || now.getMonth() == 5 || now.getMonth() == 10 )
'{ if ( now.getDate() <=21 ) { mm=monName[now.getMonth()] } else {mm=monName[now.getMonth()+1] }}
If MM >= 0 And MM < 4 Then Tm = IIf(DD <= 20, arryMon(MM), arryMon(MM + 1))
'if ( now.getMonth() == 0 || now.getMonth() == 2 || now.getMonth() == 3 )
'{ if ( now.getDate() <=20 ) { mm=monName[now.getMonth()] } else { mm=monName[now.getMonth()+1] }}
If MM = 1 Then Tm = IIf(DD <= 19, arryMon(MM), arryMon(MM + 1))
'if ( now.getMonth() == 1) { if ( now.getDate() <=19 ) { mm=monName[now.getMonth()] }
'else { mm=monName[now.getMonth()+1] }}
If MM = 11 Then Tm = IIf(DD <= 21, arryMon(MM), arryMon(MM - 11))
'if ( now.getMonth() == 11) { if ( now.getDate() <=21 ) { mm=monName[now.getMonth()] }
'else { mm=monName[now.getMonth()-11] }}
If MM = 2 Or MM = 3 Or MM = 0 Then Td = IIf(DD <= 20, DD + 10, DD + 10 - 30)
'if ( now.getMonth() == 2 || now.getMonth() == 3 || now.getMonth() == 0 )
'{if ( now.getDate() <=20 ) {mr=now.getDate()+10 } else {  mr=now.getDate()+10-30 }}
If MM = 4 Or MM = 5 Then Td = IIf(DD <= 21, DD + 10, DD + 10 - 31)
'if ( now.getMonth() == 4 || now.getMonth() == 5 )
'{ if ( now.getDate() <=21 ) { mr=now.getDate()+10 } else { mr=now.getDate()+10-31 }}
If MM > 5 And MM < 9 Then Td = IIf(DD <= 21, DD + 9, DD + 9 - 31)
'if ( now.getMonth() == 6 || now.getMonth() == 7 || now.getMonth() == 8 )
'{ if ( now.getDate() <=21 ) { mr=now.getDate()+9 } else { mr=now.getDate()+9-31 }}
If MM = 10 Or MM = 11 Then Td = IIf(DD <= 21, DD + 9, DD + 9 - 30)
'if ( now.getMonth() == 10 || now.getMonth() == 11)
'{ if ( now.getDate() <=21 ) { mr=now.getDate()+9 } else { mr=now.getDate()+9-30 }}
If MM = 9 Then Td = IIf(DD <= 22, DD + 8, DD + 8 - 31)
'if ( now.getMonth() == 9){ if ( now.getDate() <=22 ) { mr=now.getDate()+8 }
'else { mr=now.getDate()+8-30} }
If MM = 1 Then Td = IIf(DD <= 20, DD + 11, DD + 11 - 31)
'if ( now.getMonth() == 1 ){ if ( now.getDate() <=20 ) { mr=now.getDate()+11 }
'else { mr=now.getDate()+11-30  }}
IRDate2 = RR & "" & Td & "/" & MM & "/" & ii
'document.write (RR + ", " + mr + " " + MM + " " + ii)
End Function

