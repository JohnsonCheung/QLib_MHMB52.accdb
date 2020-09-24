Attribute VB_Name = "MxIde_Mthn_Mth3n"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Mth3n."

Function CMth3n$():                             CMth3n = Mth3nLn(CMthln):      End Function
Function Mth3nLn$(L):                          Mth3nLn = Mi3NtfTMth(TMthL(L)): End Function
Function Mth3nyM(M As CodeModule) As String(): Mth3nyM = Mth3nySrc(SrcM(M)):   End Function
Function Mth3nySrc(Src$()) As String()
Dim L: For Each L In MthlnySrc(Src)
    PushI Mth3nySrc, Mth3nLn(L)
Next
End Function
Function Mth3nyTMthy(N() As TMth) As String()
Dim J&: For J = 0 To UbTMth(N)
    PushI Mth3nyTMthy, Mth3nTMth(N(J))
Next
End Function
Function Mth3nTMth$(N As TMth)
With N
    If .Mthn = "" Then Exit Function
    Mth3nTMth = JnApDot(.Mthn, .ShtTy, .ShtMdy)
End With
End Function
