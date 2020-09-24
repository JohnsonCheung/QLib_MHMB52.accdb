Attribute VB_Name = "MxIde_Src_Jrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Fea_JrcyPatnPC."

Function JrcyPatnPC(Patn$, Optional PatnssAndMd$, Optional UL As eUL) As String()  ' return Sy of JrcyPatnPC, which is a srcln will jump to particular src line
Dim R As RegExp: Set R = Rx(Patn)
Dim N: For Each N In Itr(MdnyPC(PatnssAndMd))
    PushIAy JrcyPatnPC, WJrcyMdRx(N, R, UL)
Next
PushI JrcyPatnPC, "Count: " & Si(JrcyPatnPC) / IIf(UL = eULYes, 2, 1)
End Function
Private Function WJrcyMdRx(Mdn, R As RegExp, UL As eUL) As String()
Dim S$(): S = SrcMdn(Mdn)
Dim J&: For J = 0 To UB(S)
    Dim L$: L = S(J)
    Dim P As P12: P = P12Rx(L, R)
    If Not IsEmpP12(P) Then
        PushIAy WJrcyMdRx, Jrcy(Mdn, J + 1, P, L, UL)
    End If
Next
End Function

Function JrcySsub(Mdn, Lno&, Ssub, Ln$, Optional UL As eUL) As String()
JrcySsub = Jrcy(Mdn, Lno, P12Ssub(Ln, Ssub, eCasSen), Ln, UL)
End Function
Function Jrcy(Mdn, Lno&, P As P12, Ln$, Optional UL As eUL) As String()
With P
Dim Pfx$: Pfx = FmtQQ("Jmp""?:?:?:?""", Mdn, Lno, .P1, .P2) & " '"
PushI Jrcy, Pfx & Ln
If UL = eULYes Then
    Dim LnUL$: LnUL = "' " & Space(Len(Pfx) + .P1 - 3) & String(.P2 - .P1 + 1, "^")
    PushI Jrcy, LnUL
End If
End With
End Function
