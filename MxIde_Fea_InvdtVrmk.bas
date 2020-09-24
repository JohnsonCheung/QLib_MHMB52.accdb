Attribute VB_Name = "MxIde_Fea_InvdtVrmk"
Option Compare Text
Const CMod$ = "MxIde_Fea_InvdtVrmk."
Option Explicit

Sub DmpErVmkzMdn__Tst(): DmpErVmkzMdn "MxIde_Deri_Enm_Dvenm": End Sub
Sub DmpErVmkMC():        WDmp CCmpy:                          End Sub
Sub DmpErVmkPC():        WDmp CmpyPC:                         End Sub
Sub DmpErVmkzMdn(Mdn$):  WDmp CmpyMdn(Mdn):                   End Sub

Private Sub WDmp(C() As VBComponent)
Dim O$()
Dim Cmp: For Each Cmp In C
    PushIAy O, WEryzCmp(CvCmp(Cmp))
Next
DmpAy O
End Sub
Private Function WEryzCmp(C As VBComponent) As String()
Dim S$(): S = SrcCmp(C)
Dim N$: N = C.Name
PushIAy WEryzCmp, WEryzDcl(S, N)
Dim B() As Bei: B = BeiyMth(S)
Dim J&: For J = 1 To UbBei(B) - 1
    Dim EixL&: EixL = B(J - 1).Eix
    Dim BixC&: BixC = B(J).Bix
    PushIAy WEryzCmp, WEryzIx(S, EixL, BixC, N)
Next
End Function
Private Function WEryzIx(Src$(), EixL&, BixC&, Cmpn$) As String()
Dim I&: For I = EixL + 1 To BixC - 1
    Dim L$: L = Src(I)
    If IsLnVmk(L) Then
        Dim P%: P = InStr(L, "'")
        PushIAy WEryzIx, Jrcy(Cmpn, I + 1, P12(P, P + 1), L, eULYes)
    End If
Next
End Function
Private Function WEryzDcl(Src$(), Cmpn$) As String() 'Fst Mthix up, cannot have Vmk
Dim Ix&: Ix = MthixFst(Src): If Ix = -1 Then Exit Function
Dim I&: For I = Ix - 1 To 0 Step -1
    Dim L$: L = Src(I)
    If IsLnVmk(L) Then
        Dim P%: P = InStr(L, "'")
        PushIAy WEryzDcl, Jrcy(Cmpn, I + 1, P12(P, P + 1), L, eULYes)
        Exit Function
    End If
    If Trim(L) <> "" Then Exit Function
Next
End Function
