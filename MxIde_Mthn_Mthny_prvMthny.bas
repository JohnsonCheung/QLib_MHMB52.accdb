Attribute VB_Name = "MxIde_Mthn_Mthny_prvMthny"
Option Compare Text
Const CMod$ = "MxIde_Mthn_Prv."
Option Explicit


Function MthnyPrvPC() As String(): MthnyPrvPC = MthnyPrvP(CPj): End Function
Function MthnyPrvP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MthnyPrvP, MthnyPrvM(C.CodeModule)
Next
End Function

Function MthnyPrvMdn(Mdn$) As String():          MthnyPrvMdn = MthnyPrvM(Md(Mdn)): End Function
Function MthnyPrvM(M As CodeModule) As String():   MthnyPrvM = MthnyPrv(SrcM(M)):  End Function
Function MthnyPrv(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB MthnyPrv, MthnPrv(L)
Next
End Function

Function SubnyPrvM(M As CodeModule) As String(): SubnyPrvM = SubnyPrv(SrcM(M)): End Function
Function SubnyPrv(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB SubnyPrv, SubnPrv(L)
Next
End Function
Function SubnyPrvP(P As VBProject) As String():  SubnyPrvP = SubnyPrv(SrcP(P)): End Function
Function SubnyPrvPC() As String():              SubnyPrvPC = SubnyPrvP(CPj):    End Function

Function FunnyPrvM(M As CodeModule) As String(): FunnyPrvM = FunnyPrv(SrcM(M)): End Function
Function FunnyPrv(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB FunnyPrv, FunnPrv(L)
Next
End Function
Function FunnyPrvP(P As VBProject) As String():  FunnyPrvP = FunnyPrv(SrcP(P)): End Function
Function FunnyPrvPC() As String():              FunnyPrvPC = FunnyPrvP(CPj):    End Function
