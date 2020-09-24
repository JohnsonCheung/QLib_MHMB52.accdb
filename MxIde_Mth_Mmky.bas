Attribute VB_Name = "MxIde_Mth_Mmky"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Mmk."
Function MmklSrcIx$(Src$(), Mthix):            MmklSrcIx = JnCrLf(MmkySrcIx(Src, Mthix)): End Function
Function MmkySrcIx(Src$(), Mthix) As String(): MmkySrcIx = Mmky(MthySrcIx(Src, Mthix)):   End Function
Function Mmky(Mthy$()) As String()
Dim L$(): L = Contlny(Mthy)

End Function
Private Function Mmkix&(Src$(), Mthix)
If Mthix <= 0 Then Mmkix = -1: Exit Function
Dim J&, L$, I&
Mmkix = Mthix
For J = Mthix - 1 To 0 Step -1
    If Not IsLnVmkOrBlnk(Src(J)) Then
        For I = J To Mthix
            If Not IsLnBlnk(Src(I)) Then Mmkix = I: Exit Function
        Next
        ThwImposs CSub
    End If
    L = LTrim(Src(J))
    Select Case True
    Case L = ""
    Case ChrFst(L) = "'": Mmkix = J
    Case Else: Exit Function
    End Select
Next
End Function
