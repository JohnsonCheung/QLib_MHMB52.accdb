Attribute VB_Name = "MxIde_Mthn_NoPm"
#If Doc Then
'Slmb:Cml #Single-Line-Method#
'Mlm:Cml #Multi-Line-Method#
#End If
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_NoPm."
Function FunyNoPmPC() As String(): FunyNoPmPC = FunyNoPmP(CPj): End Function
Function FunyNoPmP(P As VBProject) As String()
Dim C As VBComponent:  For Each C In P.VBComponents
    PushI FunyNoPmP, FunyNoPmM(C.CodeModule)
Next
End Function
Function FunyNoPmM(M As CodeModule) As String(): FunyNoPmM = FunyNoPm(SrcM(M)): End Function
Function FunyNoPm(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLnFunNoPm(L) Then PushI FunyNoPm, MthnL(L)
Next
End Function

Function IsLnFunNoPm(L) As Boolean
If Not IsLnFun(L) Then Exit Function
IsLnFunNoPm = BetBkt(L) = ""
End Function

Private Sub B_MlmNoPmGetny()
Dim O$()
PushIAy O, Box("They should all EnsPrpOnEr")
PushIAy O, MlmNoPmGetny(SrcPC)
Brw O
End Sub
Function MlmNoPmGetny(Src$()) As String()
Dim Ix&: For Ix = 0 To UB(Src)
    PushNB MlmNoPmGetny, MlmNoPmGetn(Src, Ix)
Next
End Function
Function MlmNoPmGetn$(Src$(), Ix)
Dim L$: L = Src(Ix)
Dim N$: N = GetnNoPm(L): If N = "" Then Exit Function
If IsLnSlm(L) Then Exit Function
MlmNoPmGetn = N
End Function
