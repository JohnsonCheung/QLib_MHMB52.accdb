Attribute VB_Name = "MxIde_Mthn_Pub_mdMthnn"
Option Compare Text
Const CMod$ = "MxIde_Mthn_Pub_MMthn."
Option Explicit

Function MiMdMthnnPatnMC$(mthnPatn$): MiMdMthnnPatnMC = MiMdMthnnRxM(CMd, Rx(mthnPatn)): End Function
Function MiMdMthnnRxMC$(mthnPatn$):     MiMdMthnnRxMC = MiMdMthnnRxM(CMd, WRx):          End Function
Private Function WRx() As RegExp
Stop
End Function

Function MiMdMthnnRxM$(M As CodeModule, RxMthn As RegExp): MiMdMthnnRxM = MiMdMthnnRx(SrcM(M), Mdn(M), RxMthn): End Function
Function MiMdMthnnRx$(Src$(), Mdn$, RxMthn As RegExp)
Dim N$(): N = AwRx(MthnyPub(Src), RxMthn): If IsEmpAy(N) Then Exit Function
MiMdMthnnRx = Mdn & " " & JnSpc(N)
End Function

Function MiMdMthnnMC$():                                    MiMdMthnnMC = MiMdMthnnM(CMd):                  End Function
Function MiMdMthnnM$(M As CodeModule):                       MiMdMthnnM = MiMdMthnn(SrcM(M), Mdn(M)):       End Function
Function MiMdMthnn$(Src$(), Mdn$):                            MiMdMthnn = Mdn & " " & JnSpc(MthnyPub(Src)): End Function
Function MiMdMthnnyPC(Optional mdPatn$ = ".") As String(): MiMdMthnnyPC = MiMdMthnnyP(CPj, mdPatn):         End Function
Function MiMdMthnnyP(P As VBProject, mdPatn$) As String()
Dim R As RegExp
Set R = Rx(mdPatn)
Dim Mdny$()
    Dim C As VBComponent: For Each C In P.VBComponents
        If R.Test(C.Name) Then PushI Mdny, C.Name
    Next
    Mdny = AySrtQ(Mdny)

Dim M: For Each M In Itr(Mdny)
    PushI MiMdMthnnyP, MiMdMthnnM(P.VBComponents(M).CodeModule)
Next
End Function

Function MthnnPubM$(M As CodeModule): MthnnPubM = JnSpc(AySrtQ(MthnyPubM(M))): End Function
