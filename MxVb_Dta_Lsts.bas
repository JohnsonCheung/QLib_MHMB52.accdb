Attribute VB_Name = "MxVb_Dta_Lsts"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Lstss."
Type Lstss: N As Long: NLn As Long: Len As Long: End Type
Type Lsts: NLn As Long: Len As Long: End Type
Function Lstss(N, NLn, Len_) As Lstss
With Lstss
    .N = N
    .NLn = NLn
    .Len = Len_
End With
End Function
Function Lsts(NLn, Len_) As Lsts
With Lsts
    .NLn = NLn
    .Len = Len_
End With
End Function
Function RepLstss$(A As Lstss, Optional Lbl$ = "N-NLn-Len"): RepLstss = FmtQQ("N-NLn-Len ? ? ?", A.N, A.NLn, A.Len): End Function
Function RepLsts$(A As Lsts, Optional Lbl$ = "NLn-Len"):      RepLsts = FmtQQ("NLn-Len ? ? ?", A.NLn, A.Len):        End Function
Function LstsLy(Ly$()) As Lsts
With LstsLy
    .NLn = Si(Ly)
    .Len = LenSy(Ly)
End With
End Function
Function RepLstsLy$(Ly$())

RepLstsLy = RepLsts(LstsLy(Ly)): End Function
Function RepLstsLines$(Lines$): RepLstsLines = RepLstsLy(SplitCrLf(Lines)):               End Function
Function StsLy$(Ly$()):                StsLy = FmtQQ("NLn(?) Len(?)", Si(Ly), LenSy(Ly)): End Function

Function LenSy&(Sy$())
Dim S, O&: For Each S In Itr(Sy)
    O = O + Len(S)
Next
LenSy = O
End Function
