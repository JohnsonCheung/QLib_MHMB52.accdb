Attribute VB_Name = "MxIde_Src_SngQExmRmk"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_SngQExmRmk."
Function RmkSngQExmLin$(Ln)

End Function
Function RmkTyDfnRmkLy$(TyDfnRmkLy$())
Dim R$, O$()
Dim L: For Each L In Itr(TyDfnRmkLy)
    If ChrFst(L) = "'" Then
        Dim A$: A = LTrim(RmvFst(L))
        If ChrFst(A) = "!" Then
            PushNB O, LTrim(RmvFst(A))
        End If
    End If
Next
RmkTyDfnRmkLy = JnCrLf(O)
End Function
Function SngQExmRe() As RegExp
Static O As RegExp
If IsNothing(O) Then
End If
End Function

Function IsLnSngQExm(L) As Boolean
IsLnSngQExm = SngQExmRe.Test(L)
End Function
