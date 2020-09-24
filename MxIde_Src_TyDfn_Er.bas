Attribute VB_Name = "MxIde_Src_TyDfn_Er"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_TyDfn_Er."
Private Sub B_IsLnTyDfnEr()
Dim O$()
Dim L: For Each L In SrcP(CPj)
    If IsLnTyDfnEr(L) Then
        PushI O, L
    End If
Next
Brw O
End Sub
Function IsLnTyDfnEr(Ln) As Boolean
Dim L$: L = Ln
If TyDfnnShf(L) = "" Then Exit Function
ColonTyShf L  ' It is optional
L = Trim(L)   ' Then ! ... is must
Select Case True
Case ChrFst(L) = "!": Exit Function     '<-- It Valid line
End Select
IsLnTyDfnEr = True     '<-- It is ErLn
End Function

Function TyDfnErLnAyP() As String()
Dim L: For Each L In SrcP(CPj)
    If IsLnTyDfnEr(L) Then
        PushI TyDfnErLnAyP, L
    End If
Next
End Function

Private Sub B_TyDfnErLnAyP()
Brw TyDfnErLnAyP
End Sub
