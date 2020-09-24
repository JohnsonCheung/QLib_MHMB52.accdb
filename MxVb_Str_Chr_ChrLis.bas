Attribute VB_Name = "MxVb_Str_Chr_ChrLis"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Chr_ChrLis."

Function ChrLasLis$(S, ChrLis$, Optional C As eCas)
Dim OChrLas$: OChrLas = ChrLas(S)
If HasSsub(ChrLis, OChrLas, C) Then ChrLasLis = OChrLas
End Function
Function RmvfstChrInLis$(S, ChrLis$) ' Rmv fst chr if it is in ChrLis
If HasSsub(ChrLis, ChrFst(S)) Then
    RmvfstChrInLis = RmvFst(S)
Else
    RmvfstChrInLis = S
End If
End Function
Function RmvChrLasInLis$(S, ChrLis$) ' Rmv las chr if it is in ChrLis
If HasSsub(ChrLis, ChrLas(S)) Then
    RmvChrLasInLis = RmvLas(S)
Else
    RmvChrLasInLis = S
End If
End Function

Function TakChrFmLis$(S, ChrLis$) ' Ret fst chr if it is in ChrLis
If HasSsub(ChrLis, ChrFst(S)) Then TakChrFmLis = ChrFst(S)
End Function

Function ShfChrInLis$(OLn$, ChrLis$)
Dim C$: C = TakChrFmLis(OLn, ChrLis)
If C = "" Then Exit Function
ShfChrInLis = C
OLn = Mid(OLn, 2)
End Function
