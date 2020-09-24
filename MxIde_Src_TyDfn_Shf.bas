Attribute VB_Name = "MxIde_Src_TyDfn_Shf"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_TyDfn_Shf."

Function TyDfnnShf(OLn$)
Dim A$: A = Tm1(OLn)
If IsTyDfnn(A) Then
    TyDfnnShf = A
    OLn = RmvA1T(OLn)
End If
End Function

Function ColonTyShf$(OLn$)
':ColonTy: :Str ! #Colon-Type# it is a Tm with fst chr is : and rest is [DfnTyNm]
Dim A$: A = Tm1(OLn)
If IsTyDfnn(A) Then
    ColonTyShf = A
    OLn = RmvA1T(OLn)
End If
End Function

Function MemnShf$(OLn$)
Dim A$: A = Tm1(OLn)
If IsMemn(A) Then
    MemnShf = A
    OLn = RmvA1T(OLn)
End If
End Function
