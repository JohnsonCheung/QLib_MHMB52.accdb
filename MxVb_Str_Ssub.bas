Attribute VB_Name = "MxVb_Str_Ssub"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Ssub."

Function Right2$(S): Right2 = Right(S, 2): End Function
Function AscChrFst%(S)
AscChrFst = Asc(ChrFst(S))
End Function
Function AscChrSnd%(S)
AscChrSnd = Asc(ChrSnd(S))
End Function
Function UCasFst$(S)
UCasFst = UCase(ChrFst(S)) & RmvFst(S)
End Function
Function NSsub&(S, Ssub$, Optional C As eCas)
Dim P&: P = 1
Dim O&, L%
L = Len(Ssub)
While P > 0
    Dim J&: ThwLoopTooMuch CSub, J, 100000
    P = InStr(P, S, Ssub, C)
    If P = 0 Then NSsub = O: Exit Function
    O = O + 1
    P = P + L
Wend
End Function
Private Sub B_NSsub()
Dim A$, Ssub$
A = "aaaa":                 Ssub = "aa":  Ept = CLng(2): GoSub Tst
A = "aaaa":                 Ssub = "a":   Ept = CLng(4): GoSub Tst
A = "skfdj skldfskldf df ": Ssub = " ":   Ept = CLng(3): GoSub Tst
Exit Sub
Tst:
    Act = NSsub(A, Ssub)
    C
    Return
End Sub

Function NDot&(S):     NDot = NSsub(S, "."): End Function
Function Left2$(S):   Left2 = Left(S, 2):    End Function
Function Left3$(S):   Left3 = Left(S, 3):    End Function
Function Right3$(S): Right3 = Left(S, 3):    End Function
