Attribute VB_Name = "MxVb_Dta_AscIs"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_AscIs."
Const AscPlus% = &H2B  ' + sign
Const AscMinus% = &H2D  ' - sign
Const AscDash% = 95 '_


Function AscAyNonPrt() As Integer()
Dim J%: For J = 0 To 255
    If IsAscNonPrt(J) Then PushI AscAyNonPrt, J
Next
End Function

Function IsAscPrt(A%) As Boolean
Select Case A
Case 0, 1, 9, 10, 13, 28, 29, 30, 31, 129, 141, 143, 144, 157, 160
Case Else: IsAscPrt = True
End Select
End Function
Function IsAscNonPrt(A%) As Boolean:           IsAscNonPrt = Not IsAscPrt(A):               End Function
Function IsAscFstNmChr(A%) As Boolean:       IsAscFstNmChr = IsAscLetter(A):                End Function
Function IsAscDig(A%) As Boolean:                 IsAscDig = IsBet(A, &H30, &H39):          End Function
Function IsAscSgn(A%) As Boolean:                 IsAscSgn = A = AscPlus Or A = AscMinus:   End Function
Function IsAscDigSgn(A%) As Boolean:           IsAscDigSgn = IsAscDig(A) Or IsAscSgn(A):    End Function
Function IsAscDash(A%) As Boolean:               IsAscDash = A = 95:                        End Function
Function IsAscLCas(A%) As Boolean:               IsAscLCas = IsBet(A, 97, 122):             End Function
Function IsAscLetterOrDig(A%) As Boolean: IsAscLetterOrDig = IsAscLetter(A) Or IsAscDig(A): End Function
Function IsAscLetter(A%) As Boolean:           IsAscLetter = IsAscUCas(A) Or IsAscLCas(A):  End Function
Function IsAscNmChr(A%) As Boolean:             IsAscNmChr = IsAscLetter(A) Or IsAscDig(A): End Function
Function IsAscUCas(A%) As Boolean:               IsAscUCas = IsBet(A, 65, 90):              End Function
Function IsLetter(V$) As Boolean:                 IsLetter = IsAscLetter(Asc(V)):           End Function
Function IsDig(C$) As Boolean:                       IsDig = IsAscDig(Asc(C)):              End Function
Function IsLCas(C$):                                IsLCas = IsAscLCas(Asc(C)):             End Function
Function IsUCas(C$):                                IsUCas = IsAscUCas(Asc(C)):             End Function
Function IsWhiteChr(C$)
Select Case C
Case " ", vbTab, vbCr, vbLf: IsWhiteChr = True
End Select
End Function
Function AscAt%(C, At): AscAt = Asc(Mid(C, At, 1)): End Function
Function IsAscSpc(A%) As Boolean
Select Case A
Case 13, 10, 32, 9: IsAscSpc = True
End Select
End Function
