Attribute VB_Name = "MxIde_Dcl_Cnst_CnstStr"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Cnst_CnstStr."
Function CnststrLn$(Ln, Cnstn$) ' ret Strv of the line is a const line of name @CnstnL.  The @@Strv is assumed as the string in between of the DblQ
If IsLnHasCnstn(Ln, Cnstn) Then CnststrLn = RmvFstLas(Trim(Aft(Ln, "=")))
End Function
Function CnststrM(M As CodeModule, Cnstn$)
Dim L, O$, J%: For J = 1 To M.CountOfLines
    O = CnststrLn(M.Lines(J, 1), Cnstn)
    If O <> "" Then CnststrM = O: Exit Function
Next
End Function

Function Cnststrv(Dcl$(), Cnstn$)
Dim L, O$: For Each L In Itr(Dcl)
    O = CnststrLn(L, Cnstn)
    If O <> "" Then Cnststrv = O: Exit Function
Next
End Function
