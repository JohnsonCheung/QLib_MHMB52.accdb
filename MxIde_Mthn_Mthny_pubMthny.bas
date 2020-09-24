Attribute VB_Name = "MxIde_Mthn_Mthny_pubMthny"
Option Compare Text
Option Explicit


Function MtnyPubDashPC() As String():         MtnyPubDashPC = AwSsubDash(MthnyPubPC): End Function
Function MthnyPubNonDashPC() As String(): MthnyPubNonDashPC = AeSsubDash(MthnyPubPC): End Function

Function MthnyPubPC(Optional PatnssAndMd$, Optional SsMthPatn$) As String(): MthnyPubPC = MthnyPubP(CPj, PatnssAndMd, SsMthPatn): End Function
Function MthnyPubP(P As VBProject, Optional WhSsubssMd$, Optional mthPatnorss$) As String()
Dim R() As RegExp: R = Rxay(mthPatnorss)
Dim M$(): M = MdnyP(P, WhSsubssMd)
Dim Mdn: For Each Mdn In Itr(M)
    Dim Ny$(): Ny = MthnyPub(SrcCmp(P.VBComponents(Mdn)))
    PushIAy MthnyPubP, AwRxAyOr(Ny, R)
Next
End Function

Function MthnyPubM(M As CodeModule) As String(): MthnyPubM = MthnyPub(SrcM(M)): End Function

Function MthnyPub(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB MthnyPub, MthnPub(L)
Next
End Function

Function SubnyPubM(M As CodeModule) As String(): SubnyPubM = SubnyPub(SrcM(M)): End Function
Function SubnyPub(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB SubnyPub, SubnPub(L)
Next
End Function
Function SubnyPubP(P As VBProject) As String():  SubnyPubP = SubnyPub(SrcP(P)): End Function
Function SubnyPubPC() As String():              SubnyPubPC = SubnyPubP(CPj):    End Function
Function PrpnyPubP(P As VBProject) As String():  PrpnyPubP = PrpnyPub(SrcP(P)): End Function
Function PrpnyPubPC() As String():              PrpnyPubPC = PrpnyPubP(CPj):    End Function
Function PrpnyPub(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB PrpnyPub, PrpnPubL(L)
Next
End Function
Function FunnyPubM(M As CodeModule) As String(): FunnyPubM = FunnyPub(SrcM(M)): End Function
Function FunnyPub(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB FunnyPub, FunnPub(L)
Next
End Function
Function FunnyPubP(P As VBProject) As String():  FunnyPubP = FunnyPub(SrcP(P)): End Function
Function FunnyPubPC() As String():              FunnyPubPC = FunnyPubP(CPj):    End Function

Sub VcMthnyPubDash(): Vc AwSsubDash(MthnyPubPC): End Sub
Sub VcFunnyPubDash(): Vc AwSsubDash(FunnyPubPC): End Sub
