Attribute VB_Name = "MxIde_Mthln_MthTy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln_MthTy."


Function MthTyLn$(Ln): MthTyLn = PfxPfxySpc(RmvMdy(Ln), MthTyy): End Function

Private Sub B_MthTyLn()
Dim O$(), L
For Each L In SrcPC
    Dim T$: T = ShtMthTy(MthTyLn(L))
    If T <> "" Then Push O, T & " " & L
Next
BrwAy O
End Sub

Function IsMthTy(S) As Boolean: IsMthTy = HasEle(MthTyy, S): End Function
Function IsPrpTy(S) As Boolean: IsPrpTy = HasEle(PrpTyy, S): End Function

Function IsShtMthTyRet(ShtMthTy$) As Boolean
Select Case True
Case "Fun", "Get": IsShtMthTyRet = True
End Select
End Function
Function IsMthTyRet(Mtht$) As Boolean
Select Case True
Case "Function", "Property Get": IsMthTyRet = True
End Select
End Function
