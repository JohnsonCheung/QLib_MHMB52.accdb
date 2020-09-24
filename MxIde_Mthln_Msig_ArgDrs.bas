Attribute VB_Name = "MxIde_Mthln_Msig_ArgDrs"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_MSig_ArgDrs."
Public Const FfTArg$ = "Mthn No Nm IsOpt IsByVal IsPmAy IsAy Tyc AsTy DftVal"

Private Sub B_DrsTArgPC():                             BrwDrs DrsTArgPC:                End Sub
Private Sub B_DrsTArgMC():                             BrwDrs DrsTArgMC:                End Sub
Function DrsTArgM(M As CodeModule) As Drs:  DrsTArgM = DrsFf(FfTArg, X_Dy(MthlnyM(M))): End Function
Function DrsTArgP(P As VBProject) As Drs:   DrsTArgP = DrsFf(FfTArg, X_Dy(MthlnyP(P))): End Function
Function DrsTArgPC() As Drs:               DrsTArgPC = DrsTArgP(CPj):                   End Function
Function DrsTArgMC() As Drs:               DrsTArgMC = DrsTArgM(CMd):                   End Function

Function DrsTArgVC() As Drs:                 DrsTArgVC = W2Drs(MthlnyVC):                End Function
Private Function W2Drs(MthlnySrc$()) As Drs:     W2Drs = DrsFf(FfTArg, X_Dy(MthlnySrc)): End Function

Private Function X_Dy(MthlnySrc$()) As Variant()
Dim L: For Each L In Itr(MthlnySrc)
    PushIAy X_Dy, W2Dy(L)
Next
End Function
Private Function W2Dy(Mthln) As Variant()
Dim Pm$: Pm = BetBkt(Mthln)
Dim A$(): A = AmTrim(SplitCmaSpc(Pm))
Dim N$: N = MthnL(Mthln)
Dim Arg$, Dy(), ArgNo%: For ArgNo = 1 To Si(A)
    Arg = A(ArgNo - 1)
    PushI W2Dy, W2Dr(Arg, ArgNo, N)
Next
End Function
Private Function W2Dr(Arg$, ArgNo%, Mthn$) As Variant()
Dim A As TArg: A = TArgArg(Arg)
Dim M As eArgm: M = A.Argm
Dim IsOpt   As Boolean:   IsOpt = M = eArgmRefOpt Or M = eArgmValOpt
Dim IsByVal As Boolean: IsByVal = M = eArgmVal Or M = eArgmValOpt
Dim IsPmAy  As Boolean:  IsPmAy = M = eArgmAp
Dim Nm$:                     Nm = A.Argn
Dim Tyc$:               Tyc = A.Vt.Tyc
Dim IsAy    As Boolean:    IsAy = A.Vt.IsAy
Dim Tyn$:                   Tyn = A.Vt.Tyn
W2Dr = Array(Mthn, ArgNo, Nm, IsOpt, IsByVal, IsPmAy, IsAy, Tyc, Tyn, A.Dft)
End Function
