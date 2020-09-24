Attribute VB_Name = "MxIde_Mthln_Fun"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln_Fun."
Function Funlny(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLnFun(L) Then
        PushI Funlny, L
    End If
Next
End Function
Function FunlnyM(M As CodeModule) As String():  FunlnyM = Funlny(SrcM(M)): End Function
Function FunlnyMC() As String():               FunlnyMC = FunlnyM(CMd):    End Function

Function FunlnyPC() As String(): FunlnyPC = FunlnyP(CPj): End Function
Function FunlnyP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy FunlnyP, FunlnyM(C.CodeModule)
Next
End Function

Function FunlnyPubPC() As String(): FunlnyPubPC = FunlnyPubP(CPj): End Function
Function FunlnyPubP(P As VBProject) As String()
Dim L: For Each L In MthlnyP(P)
    If W3IsLnPubFun(L) Then PushI FunlnyPubP, L
Next
End Function
Private Function B_W3IsLnPubFun()
Dim O$()
    Dim L: For Each L In SrcPC
        If W3IsLnPubFun(L) Then PushI O, L
    Next
VcAy O
End Function
Private Function W3IsLnPubFun(DfnLn) As Boolean
Select Case Mdy(DfnLn)
Case "", "Public"
    W3IsLnPubFun = HasPfxSpc(RmvPfxSpc(DfnLn, "Public"), "Function")
End Select
End Function
