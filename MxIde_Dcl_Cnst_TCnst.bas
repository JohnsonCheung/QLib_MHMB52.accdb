Attribute VB_Name = "MxIde_Dcl_Cnst_TCnst"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Cnst_TCnst."
Type TCnst: Mdn As String: Cnstn As String: Tycn As String: CnstRem As String: End Type
Function TCnstyM(M As CodeModule, Optional InlCModv As Boolean) As TCnst()
Dim Src$(): Src = SrcM(M)
Dim Ix&, S As S12: For Ix = 0 To UB(Src)
    Dim L$: L = ContlnIx(Src, Ix)
    S.S1 = CnstnL(L)
    If S.S1 = "" Then GoTo Nxt
    If Not InlCModv Then If S.S1 = "CMod" Then GoTo Nxt
    S.S2 = Aft(L, "=")
    Stop 'PushS12 S12yCnstM, S
Nxt:
Next
End Function
Private Sub B_TCnstyM()
Stop 'BrwTCnsty TCnstyCnst(CMd)
End Sub
