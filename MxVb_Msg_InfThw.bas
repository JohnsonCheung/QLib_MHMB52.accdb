Attribute VB_Name = "MxVb_Msg_InfThw"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Run_PM_Thw_Inf1."
Private Enum eHaltTy: ePgmEr: ePmEr: eLgcEr: eImposs: eLoopTooMuch: eUsrInf: eUsrWarn: End Enum ' Deriving(Str Txt)
Const EnmttmlHaltTy$ = "[Program Error] [Parameter Error] [Logic Error] [Impossible to reach here] [Looping too much] [User information] [User warning]"
Const EnmqssHaltTy$ = "ePgmEr ePmEr eLgcEr eImposs eLoopTooMuch eUsrInf eUsrWarn"
Private Function EnmsHaltTy$(E As eHaltTy)
Static Sy$(): If Si(Sy) = 0 Then Sy = NyQss(EnmqssHaltTy)
EnmsHaltTy = Sy(E)
End Function
Private Function EnmsyHaltTy() As String()
Static Sy$(): If Si(Sy) = 0 Then Sy = NyQss(EnmqssHaltTy)
EnmsyHaltTy = Sy
End Function
Private Function EnmvHaltTy(S) As eHaltTy: EnmvHaltTy = IxEle(EnmsyHaltTy, S): End Function
Private Function EnmtHaltTy$(E As eHaltTy)
Static T$(): If Si(T) = 0 Then T = Tmy(EnmttmlHaltTy)
EnmtHaltTy = EleMsg(T, E)
End Function

Sub ThwLoopTooMuch(Fun$, OCnt, Optional Max = 10000)
DoEvents: OCnt = OCnt + 1: If OCnt > Max Then RaiseMsg WMsgThw(eLgcEr, Fun, "Looping too much", Av("Cnt", OCnt))
End Sub
Private Sub B_Thw():                   Thw "SF", "AF", "A B C", 1, 2, 3:      End Sub
Sub Thw(Fun$, Msg$, ParamArray Nap()): Dim Av(): Av = Nap: WThw Fun, Msg, Av: End Sub
Sub ThwTrue(IfTrue As Boolean, Fun$, Msg$, ParamArray Nap())
If IfTrue = True Then
    Dim Av(): Av = Nap: WThw Fun, Msg, Av
End If
End Sub
Sub ThwFalse(IfFalse As Boolean, Fun$, Msg$, ParamArray Nap())
If IfFalse = False Then
    Dim Av(): Av = Nap: WThw Fun, Msg, Av
End If
End Sub

Sub ThwImposs(Fun$, Optional Reason$):             Thw Fun, "Impossible to reach here: " & Reason:                  End Sub
Sub ThwImpossNap(Fun$, Reason$, ParamArray Nap()): Thw Fun, "Impossible to reach here: " & Reason:                  End Sub
Sub ThwPm(Fun$, Msgln$, ParamArray Nap()):         Dim Nav(): Nav = Nap: RaiseMsg WMsgThw(ePmEr, Fun, Msgln, Nav):  End Sub
Sub ThwLgc(Fun$, Msgln$, ParamArray Nap()):        Dim Nav(): Nav = Nap: RaiseMsg WMsgThw(eLgcEr, Fun, Msgln, Nav): End Sub
Sub ThwEnm(Fun$, EnmvOrEnms, Enmqss$)
Dim U%: U = UB(SplitSpc(Enmqss)) - 1
Stop 'Dim Nn$: Nn = NnQtp(Enmm)
Stop 'Thw Fun, "Invalid Enms/Enmv", "EnmnLn EnmvOrEnms EnmUB EnmNnVdt", N, EnmvOrEnms, U, Nn
End Sub
Sub InfNCmp(Fun$)
Debug.Print Fun; " NCmp="; CPj.VBComponents.Count
End Sub
Sub InfCnt(OCnt%, Optional Stp% = 10)
DoEvents
If OCnt Mod Stp = 0 Then Debug.Print OCnt;: If OCnt Mod Stp * 20 Then Debug.Print
OCnt = OCnt + 1
End Sub
Sub Inf(Fun$, Msg$, ParamArray Nap()):                                Dim Nav(): Nav = Nap: D MsgyFMNav(Fun, Msg, Nav):       End Sub
Sub Infln(Fun$, Msg$, ParamArray Nap()):                              Dim Nav(): Nav = Nap: D MsgFMNav(Fun, Msg, Nav):        End Sub
Private Function WMsgThw$(T As eHaltTy, Fun$, Msg$, Nav()): WMsgThw = Boxl(EnmtHaltTy(T)) & JnCrLf(MsgyFMNav(Fun, Msg, Nav)): End Function

Private Sub WThw(Fun$, Msg$, Nav())
Static IsInThw As Boolean
If IsInThw Then Raise "Thw is called recurively....."
IsInThw = True
RaiseMsg WMsgThw(ePgmEr, Fun, Msg, Nav)
IsInThw = False
End Sub
Sub RaiseLoopTooMuch(Fun$, OCnt&, Optional Max& = 1000)
OCnt = OCnt + 1: If OCnt > Max Then RaiseMsg Fun & ": Looping too much"
End Sub
