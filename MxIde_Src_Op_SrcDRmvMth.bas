Attribute VB_Name = "MxIde_Src_Op_SrcDRmvMth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Op_Rmv_Src."

Function SrcRplMthn(Src$(), Mthn$, Newl$, Optional ShtMthTy$) As String() ' Return a new @@Src after removing @Mthn+@ShtMthTy and inserting @Newl
SrcRplMthn = AyRplBeiy(Src, SplitCrLf(Newl), BeiyMthn(Src, Mthn, ShtMthTy))
End Function

Function SrcDltMthny(Src$(), Mthny$()) As String()
Dim O$(): O = Src
Dim N: For Each N In Itr(Mthny)
    O = SrcDltMthn(O, N)
Next
SrcDltMthny = O
End Function
Function SrcDltMthn(Src$(), Mthn) As String()
Dim Bei As Bei
Bei = BeiMthn(Src, Mthn)
If IsEmpBei(Bei) Then Debug.Print FmtQQ("SrcDltMthn: Given Mthn not found in Src[? lines] Mthn[?]", Si(Src), Mthn): Exit Function
Dim IsPrp As Boolean: IsPrp = IsLnPrp(Src(Bei.Bix))
Dim O$(): O = AeBei(Src, Bei)
Debug.Print FmtQQ("SrcDltMthn: Mth is DELETED Src[?/? lines bef/aft] Mthn[?]", Si(Src), Si(O), Mthn)
If Not IsPrp Then
    SrcDltMthn = O
    Exit Function
End If
Bei = BeiMthn(Src, Mthn)
If IsEmpBei(Bei) Then
    Debug.Print FmtQQ("SrcDltMthn: Mthn is non pair prp, snd Mth is not found. Src[? lines aft fst dlt] Mthn[?]", Si(O), Mthn)
    SrcDltMthn = O
    Exit Function
End If
SrcDltMthn = AeBei(O, Bei)
Debug.Print FmtQQ("DltMth: Snd prp mth is DELETED. Src[? lines after 2 dlt] Mthn[?]", Si(O), Mthn)
End Function
