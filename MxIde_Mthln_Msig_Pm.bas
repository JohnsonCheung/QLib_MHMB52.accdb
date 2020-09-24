Attribute VB_Name = "MxIde_Mthln_Msig_Pm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_MSig_Pm."


Function Mthpmy(MthlnySrc$()) As String()
Dim Mthln: For Each Mthln In Itr(MthlnySrc)
    PushI Mthpmy, BetBkt(Mthln)
Next
End Function

Function ShtMthpm$(Mthpm)
Dim Argy$(): Argy = AmLTrim(SplitCma(Mthpm))
Dim O$()
Dim Arg: For Each Arg In Itr(Argy)
    PushI O, ShtArg(Arg)
Next
ShtMthpm = JnSpc(O)
End Function
