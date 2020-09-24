Attribute VB_Name = "MxIde_Src_LnContRmkLn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_LnContRmkLn."

Function IsContRmkLn(Ln) As Boolean
Select Case True
Case Not ChrFst(LTrim(Ln)) = "'": Exit Function
Case Not ChrLas(Ln) = "_": Exit Function
End Select
IsContRmkLn = True
End Function

Private Sub B_HasContRmkLnP()
MsgBox HasContRmkLnP
End Sub

Function HasContRmkLnP() As Boolean
HasContRmkLnP = HasContRmkLn(SrcP(CPj))
End Function

Function HasContRmkLn(Src$()) As Boolean
Dim L: For Each L In Itr(Src)
    If IsContRmkLn(L) Then HasContRmkLn = True: Exit Function
Next
End Function
