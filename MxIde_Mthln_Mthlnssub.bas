Attribute VB_Name = "MxIde_Mthln_Mthlnssub"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln_Mthlnssub."

Function ShfTyc$(OLn$): ShfTyc = ShfChrInLis(OLn, LisTyc): End Function
Function TakTyc$(S):    TakTyc = TakChrFmLis(S, LisTyc):   End Function
Function MthRetcn$(Contln)
Dim M As TMth: M = TMthL(Contln)
If M.Mthn = "" Then Exit Function
Select Case M.ShtTy
Case "Get", "Fun"
Case Else: Exit Function
End Select
With S123Bkt(Contln)
    Dim C$: C = ChrLas(.S1)
    Select Case True
    Case IsTyc(C): MthRetcn = C
    Case HasPfx(.S3, " As "): MthRetcn = Trim(BefOrAll(RmvPfx(.S3, " As "), "'"))
    Case Else: MthRetcn$ = "Var"
    End Select
End With
End Function
Function ShfMtht$(OLn$)
Dim O$: O = TakMtht(OLn$)
If O = "" Then Exit Function
ShfMtht = O
OLn = LTrim(RmvPfx(OLn, O))
End Function

Function IsTyc(A) As Boolean
If Len(A) <> 1 Then Exit Function
IsTyc = HasSsub(LisTyc, A)
End Function
Function TycLn$(Ln)
If IsLnFun(Ln) Then TycLn = TakTyc(RmvMth(Ln))
End Function
