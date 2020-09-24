Attribute VB_Name = "MxVb_Str_Has"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Has."
Enum eCas: eCasIgn: eCasSen: End Enum 'Deriving(Str Val Txt)
Public Const EnmqssCas$ = "eCas? Ign Sen"
Function EnmsyCas() As String()
Static S$(): If Si(S) = 0 Then S = NyQss(EnmqssCas)
EnmsyCas = S
End Function
Function EnmsCas$(E As eCas):       EnmsCas = EleMsg(EnmsyCas, E): End Function
Function HasDot(S) As Boolean:       HasDot = HasSsub(S, "."):     End Function
Function HasQuoSng(S) As Boolean: HasQuoSng = InStr(S, vbQuoSng):  End Function
Function HasQuoDbl(S) As Boolean: HasQuoDbl = InStr(S, vbQuoDbl):  End Function

Private Sub B_RmvVstr()
Dim S$
GoSub T1
GoSub T2
GoSub YY
Exit Sub
T1:
    S = "("""""""")"
    Ept = "("""""""")"
    GoTo Tst
T2:
    S = "For Each I In AwSsub(AwSsub(SrcP(CPj), ""'"")"
    Ept = "For Each I In AwSsub(AwSsub(SrcP(CPj), """")"
    GoTo Tst
Tst:
    Act = RmvVstr(S)
    C
    Return
YY:
    Dim S12y() As S12
    Dim L: For Each L In SrcP(CPj)
        If Not IsLnVmk(L) Then
            If HasQuoDbl(L) Then
                PushS12 S12y, S12(L, RmvVstr(L))
            End If
        End If
    Next
    BrwS12y S12y
    Return
End Sub

Function RmvBetDblQ$(S)
Dim P&: P = InStr(S, vbQuoDbl)
Dim O$: O = S
While P > 0
    Dim J%: J = J + 1: If J > 10000 Then Stop
    Dim P1&: P1 = InStr(P + 1, O, vbQuoDbl): If P1 = 0 Then Stop
    O = Left(O, P) & Mid(O, P1)
    P = InStr(P + 2, O, vbQuoDbl)
Wend
RmvBetDblQ = O
End Function

Private Sub B_NDblQUo()
GoSub T1
Exit Sub
Dim S$
T1:
    S = "JrclnLno = FmtQQ(""Jmp""""?"
    Ept = 3
    GoTo Tst
Tst:
    Act = NDblQuo(S)
    C
    Return
End Sub
Function NDblQuo%(S): NDblQuo = NSsub(S, vbQuoDbl): End Function

Function HasSsubDash(S):                                                       HasSsubDash = HasSsub(S, "_"):                        End Function
Function HasSsub(S, Ssub, Optional C As eCas, Optional fmPos% = 1) As Boolean:     HasSsub = InStr(fmPos, S, Ssub, VbCprMth(C)) > 0: End Function
Function HasSsubyAnd(S, Ssuby$(), Optional C As eCas) As Boolean
Dim Ssub: For Each Ssub In Itr(Ssuby)
    If Not HasSsub(S, Ssub, C) Then Exit Function
Next
HasSsubyAnd = True
End Function
Function HasSsubyOr(S, Ssuby$(), Optional C As eCas) As Boolean
Dim Ssub: For Each Ssub In Itr(Ssuby)
    If HasSsub(S, Ssub, C) Then HasSsubyOr = True: Exit Function
Next
End Function
Function HasCr(S) As Boolean:       HasCr = HasSsub(S, vbCr):                    End Function
Function HasLf(S) As Boolean:       HasLf = HasSsub(S, vbLf):                    End Function
Function HasCrLf(S) As Boolean:   HasCrLf = HasSsub(S, vbCrLf):                  End Function
Function HasHyp(S) As Boolean:     HasHyp = HasSsub(S, "-"):                     End Function
Function HasPound(S) As Boolean: HasPound = InStr(S, "#") > 0:                   End Function
Function HasSpc(S) As Boolean:     HasSpc = InStr(S, " ") > 0:                   End Function
Function HasBktSq(S) As Boolean: HasBktSq = ChrFst(S) = "[" And ChrLas(S) = "]": End Function
Function NoLf(S) As Boolean:         NoLf = Not HasLf(S):                        End Function
Function HasChrLis(S, ChrLis$, Optional Cpr As VbCompareMethod) As Boolean
Dim J%
For J = 1 To Len(ChrLis)
    If HasSsub(S, Mid(ChrLis, J, 1), Cpr) Then HasChrLis = True: Exit Function
Next
End Function
Function HasVbar(S) As Boolean
HasVbar = HasSsub(S, "|")
End Function
