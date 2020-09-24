Attribute VB_Name = "MxVb_Msg_VbMsg"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Nav_FmtNav."
Function MsgyNNAv(NN$, Av()) As String(): MsgyNNAv = MsgyNyAv(Tmy(NN), Av): End Function
Function MsgyNyAv(Ny$(), Av()) As String()
If Si(Ny) <> Si(Av) Then RaiseQQ "MsgyNyAv: Given Ny and Av is invalid: NySi[?] AvSi[?]", Si(Ny), Si(Av)
Dim N$(): N = AmAli(Ny)
Dim J%: For J = 0 To UB(N)
    PushIAy MsgyNyAv, MsglyNmV(N(J), Av(J))
Next
End Function

Function MsgyFMNap(Fun$, M$, ParamArray Nap()) As String() '#FMNap:Fun$-Msg$-Nap():Cml#
Dim Nav(): Nav = Nap
MsgyFMNap = MsgyFMNav(Fun, M, Nav)
End Function
Function MsgyFMNav(Fun$, Msg$, Nav()) As String()
If Si(Nav) = 0 Then MsgyFMNav = MsgyNNAv("Fun Msg", Av(Fun, Msg)): Exit Function
Dim NN1$: NN1 = "Fun Msg " & Nav(0)
Dim Av1(): Av1 = AyAdd(Array(Fun, Msg), AeFst(Nav))
MsgyFMNav = MsgyNNAv(NN1, Av1)
End Function

Function MsgyNNAp(NN$, ParamArray Ap()) As String()
Dim Av(): Av = Ap: MsgyNNAp = MsgyNNAv(NN, Av)
End Function
Function MsgyNav(Nav()) As String()
Dim NN$: NN = Nav(0)
Dim Av(): Av = AeFst(Nav)
MsgyNav = MsgyNNAv(NN, Av)
End Function
Private Sub B_MsgyNav(): D MsgyNav(Av("aa bb", 1, 2)): End Sub

Function MsgNNAv$(NN$, Av())
Dim Ny$(): Ny = Tmy(NN)
If Si(Ny) <> Si(Av) Then RaiseQQ "MsgyNNAv: Given Ny and Av is invalid: NySi[?] AvSi[?]", Si(Ny), Si(Av)
Dim O$()
Dim J%: For J = 0 To UB(Ny)
    PushI O, "-" & QuoTm(Ny(J)) & " " & QuoTm(Av(J))
Next
MsgNNAv = JnSpc(O)
End Function
Function MsgFMNap$(Fun$, Msg$, ParamArray Nap()): Dim Nav(): Nav = Nap: MsgFMNap = MsgFMNav(Fun, Msg, Nav): End Function
Function MsgFMNav$(Fun$, Msg$, Nav()):
MsgFMNav = Now & " " & Msg & " (@" & Fun & ") " & MsgNav(Nav)
End Function
Function MsgNav$(Nav())
If Si(Nav) = 0 Then Exit Function
Dim NN$: NN = Nav(0)
Dim Av(): Av = AeFst(Nav)
MsgNav = MsgNNAv(NN, Av)
End Function
