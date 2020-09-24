Attribute VB_Name = "MxIde_Mth_Slm_zIntl_AliSlmSrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Slmb_SlmbAli."
Private Sub B_SrcoptSlm()
GoSub T1
GoSub ZZ
Exit Sub
Dim O() As Cpmd, M As Lyopt, S$(), C As VBComponent, S12y() As S12
Dim Src$()
Dim Act As Lyopt, Ept As Lyopt
ZZ:
    For Each C In CPj.VBComponents
        Src = SrcCmp(C)
        M = SrcoptSlm(Src)
        If M.Som Then
            PushCpmd O, Cpmd(C.Name, Src, M.Ly)
        End If
    Next
    CprCpmd O
    Return
ZZ1:
    For Each C In CPj.VBComponents
        S = SrcM(C.CodeModule)
        M = SrcoptSlm(S)
        If M.Som Then
            PushS12 S12y, S12(JnCrLf(S), JnCrLf(M.Ly))
        End If
    Next
    BrwLsy LsyS12yFc(S12y)
    Return
T1:
    Src = SrcMdn("MxXls_FxWsOp")
    Ept = SomLy(Src)
    GoTo Tst
Tst:
    Act = SrcoptSlm(Src)
    If Not IsEqLyopt(Ept, Act) Then Stop
    Return
End Sub
    
Function SrcoptSlm(Src$()) As Lyopt: SrcoptSlm = LyoptOldNew(LyEndTrim(Src), LyEndTrim(WSrcAli(Src))): End Function

Private Function WSrcAli(Src$()) As String()
Dim O$(): O = Src
Dim Beiy() As Bei: Beiy = BeiySlm(Src)
Dim SLmby(): SLmby = AyyBeiy(Src, Beiy)
Dim J&: For J = 0 To UB(SLmby)
    With SlmboptAli(CvSy(SLmby(J)))
        If .Som Then
            O = SyRpl(O, .Ly, Beiy(J).Bix)
        End If
    End With
Next
WSrcAli = O
End Function


