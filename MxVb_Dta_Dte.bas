Attribute VB_Name = "MxVb_Dta_Dte"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Dte."
Public NoStamp As Boolean
Sub Stamp(S)
If Not NoStamp Then Debug.Print StrNow; " "; S
End Sub


Sub SampFunA()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        'Debug.Print I
    Next
Next
End Sub
Sub SampFunB()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        'Debug.Print I
    Next
Next
End Sub

Function CYY() As Byte: CYY = CYr - 2000: End Function ' current year in 20YY
Function CYr%():        CYr = Year(Now):  End Function ' current year in byte
Function CMM() As Byte: CMM = Month(Now): End Function   ' current month in byte
Function CDD() As Byte: CDD = Day(Now):   End Function

Function DteFstNxtMth(D As Date) As Date: DteFstNxtMth = DteFst(DteNxtMth(D)): End Function

Function IsDteLas(D As Date) As Boolean:  IsDteLas = DtePrv(DteFstNxtMth(D)) = D:                End Function
Function RemDays(D As Date) As Byte:       RemDays = NDay(D) - Day(D):                           End Function
Function NDay(D As Date) As Byte:             NDay = Day(DtePrv(DteFstNxtMth(D))):               End Function
Function DteNxtMth(D As Date) As Date:   DteNxtMth = DateAdd("M", 1, D):                         End Function
Function DtePrvMth(D As Date) As Date:   DtePrvMth = DateAdd("M", -1, D):                        End Function
Function DteFst(D As Date) As Date:         DteFst = DateSerial(Year(D), Month(D), 1):           End Function
Function DteLas(D As Date) As Date:         DteLas = DtePrv(DteFst(DteFstNxtMth(D))):            End Function
Function DtePrv(D As Date) As Date:         DtePrv = DateAdd("D", -1, D):                        End Function
Function StrYymYmM$(D As Date):          StrYymYmM = Right(Year(D), 2) & Format(Month(D), "00"): End Function

Function DteFstYM(M As Ym) As Date: DteFstYM = DateSerial(2000 + M.Y, M.M, 1):  End Function
Function DteLasYM(M As Ym) As Date: DteLasYM = DteFstNxtMth(DteFstYM(M)):       End Function
Function YmDte(D As Date) As Ym:       YmDte = Ym(YYzDte(D) - 2000, MMzDte(D)): End Function
Function YYzDte(D As Date) As Byte:   YYzDte = Year(D) - 2000:                  End Function
Function MMzDte(D As Date) As Byte:   MMzDte = Month(D):                        End Function

Function YmNxt(M As Ym) As Ym
If M.M = 12 Then
    YmNxt = Ym(M.Y + 1, 1)
Else
    YmNxt = Ym(M.Y, M.M + 1)
End If
End Function
Function YmPrv(M As Ym) As Ym
If M.M = 1 Then
    YmPrv = Ym(M.Y - 1, 12)
Else
    YmPrv = Ym(M.Y, M.M - 1)
End If
End Function

Function IsMnth(M As Byte): IsMnth = IsBet(M, 1, 23):                                                                   End Function
Sub ChkIsMnth(M As Byte):            ThwTrue IsMnth(M), "ChkIsM", FmtQQ("M should be between 1 and 12, but now[?]", M): End Sub

Function FactorRemDays!(D As Date)
Dim R As Byte: R = RemDays(D): If R = 0 Then Exit Function
FactorRemDays = R / NDay(D)
End Function
Private Sub B_NyMonthM3SpcY4()
Dim J%, A$(): A = NyMonthM3SpcY4(Ym(19, 12))
For J = 0 To 11
    Debug.Print A(J)
Next
End Sub
Function DteyYm(M As Ym, Optional NMth% = 15) As Date()
Dim D As Date: D = DteFstYM(M)
Dim J%: For J = 0 To NMth - 1
    PushI DteyYm, D
    D = DteNxtMth(D)
Next
End Function
Function NyMonthM3SpcY4(M As Ym, Optional NMth% = 15) As String()
Dim D() As Date: D = DteyYm(M, NMth)
Dim J%: For J = 0 To NMth - 1
    PushS NyMonthM3SpcY4, Format(D(J), "MMM YYYY")
Next
End Function
