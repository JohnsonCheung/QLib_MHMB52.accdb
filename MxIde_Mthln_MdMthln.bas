Attribute VB_Name = "MxIde_Mthln_MdMthln"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln_MdMthln."
Function Mi2yMdMthlnyPC() As String()
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushI Mi2yMdMthlnyPC, Mi2yMdMthlnyM(C.CodeModule)
Next
End Function
Function Mi2yMdMthlnyM(M As CodeModule) As String()
Dim P$: P = Mdn(M) & " "
Mi2yMdMthlnyM = AmAddPfx(MthlnyM(M), P)
End Function
