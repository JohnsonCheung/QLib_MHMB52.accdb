Attribute VB_Name = "MxIde_Dcl_Cnst_CnstlnOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Cnst_CnstlnOp."

Sub EnsCnstln(M As CodeModule, Cnstln$, Optional AftCnstn$)
Dim Lno&: Lno = CnstlnoN(M, CnstnL(Cnstln))
If Lno > 0 Then
    EnsCdln M, Lno, Cnstln
    Exit Sub
End If
InsCnstln M, Cnstln, AftCnstn
End Sub

Sub InsCnstln(M As CodeModule, Cnstln$, Optional AftCnstn$)
Dim Lno&
    Lno = CnstlnoL(M, AftCnstn): If Lno <> 0 Then Lno = Lno + 1
    If Lno = 0 Then Lno = LnoAftOptAndImp(M)
InsCdl M, Lno, Cnstln
End Sub

Sub DltCnstln(M As CodeModule, Cnstln$): DltDclln M, Cnstln: End Sub
