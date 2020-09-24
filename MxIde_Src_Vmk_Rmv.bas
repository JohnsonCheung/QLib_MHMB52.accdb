Attribute VB_Name = "MxIde_Src_Vmk_Rmv"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Vmk_Rmv."


Function SrcRmvVmkMC() As String():               SrcRmvVmkMC = SrcRmvVmkM(CMd):    End Function
Function SrcRmvVmkPC() As String():               SrcRmvVmkPC = SrcRmvVmkP(CPj):    End Function
Function SrcRmvVmkM(M As CodeModule) As String():  SrcRmvVmkM = SrcRmvVmk(SrcM(M)): End Function
Function SrcRmvVmkP(P As VBProject) As String():   SrcRmvVmkP = SrcRmvVmk(SrcP(P)): End Function

Private Sub B_SrcRmvVmk()
BrwAy SrcRmvVmk(SrcPC)
End Sub
Function SrcRmvVmk(Src$()) As String()
Dim L: For Each L In Itr(Contlny(Src))
    PushNB SrcRmvVmk, RmvVmk(L)
Next
End Function
