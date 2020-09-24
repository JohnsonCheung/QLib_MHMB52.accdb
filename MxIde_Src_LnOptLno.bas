Attribute VB_Name = "MxIde_Src_LnOptLno"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_LnOptLno."
Private Sub B_LnoAftOptAndImp()
GoSub T1
Exit Sub
Dim M As CodeModule
T1:
    Set M = CMd
    Ept = 3&
    GoTo Tst
Tst:
    Act = LnoAftOptAndImp(M)
    C
    Return
End Sub
Function LnoAftOptAndImp&(M As CodeModule)
Dim N%: N = M.CountOfDeclarationLines
Dim J%: For J = 1 To N
    Dim L$: L = M.Lines(J, 1)
    If Not WIsLnOptOrImpOrBlnk(L) Then LnoAftOptAndImp = J: Exit Function
Next
LnoAftOptAndImp = N + 1
End Function
Private Function WIsLnOptOrImpOrBlnk(L) As Boolean
Select Case True
Case IsLnOpt(L), IsLnImp(L), IsLnBlnk(L): WIsLnOptOrImpOrBlnk = True
End Select
End Function
