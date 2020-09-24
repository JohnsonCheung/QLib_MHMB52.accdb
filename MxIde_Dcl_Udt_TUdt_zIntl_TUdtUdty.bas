Attribute VB_Name = "MxIde_Dcl_Udt_TUdt_zIntl_TUdtUdty"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_TUdt_TUdtUdty."
Function TUdtUdty(Udty$(), Mdn$) As TUdt
Dim S$(): S = StmtySrc(Udty)
Dim O As TUdt
O.Mdn = Mdn
WSet_5IsGen O, Udty
WSet_IsPrvAndUdtn O, S(0)
WSet_Mbr O, S
TUdtUdty = O
End Function
Private Sub WSet_5IsGen(O As TUdt, Udty$()): WSet_5IsGen_ByBetDeriving O, WBetDeriving(Udty): End Sub
Private Sub WBetDeriving__Tst()
GoSub T1
Exit Sub
Dim Udt$()
T1: PushI Udt, "Type Swl: Swn As String: Op As String: Tml() As String: End Type ' Deriving(Ctor Ay)"
    Ept = "Ctor Ay"
    GoTo Tst
Tst:
    Act = WBetDeriving(Udt)
    C
    Return
End Sub
Private Function WBetDeriving$(Udt$())
WBetDeriving = Nmbktv(Udt(0), "Deriving")
If WBetDeriving <> "" Then Exit Function
If Si(Udt) = 1 Then Exit Function
WBetDeriving = Nmbktv(EleLas(Udt), "Deriving")
End Function
Private Sub WSet_5IsGen_ByBetDeriving(O As TUdt, Bet$)
Const CSub$ = CMod & "WSet_5IsGen_ByBetDeriving"
Dim N: For Each N In ItrSS(Bet)
    Select Case N
    Case "Ay": O.GenAy = True
    Case "Ctor": O.GenCtor = True
    Case "Opt": O.GenOpt = True
    Case "AyAdd": O.GenAdd = True: O.GenAy = True
    Case "PushAy": O.GenPushAy = True: O.GenAy = True
    Case Else: Thw CSub, "The * in Deriving(*) has invalid value.  Valid value are (Ay Ctor Opt)", "Bet", Bet
    End Select
Next
End Sub
Private Sub WSet_IsPrvAndUdtn(O As TUdt, StmtFst$) 'Set IsPrv & UdtnLn
Const CSub$ = CMod & "WSet_IsPrvAndUdtn"
Dim L$: L = StmtFst
O.IsPrv = ShfMdy(L) = "Private"
If Not IsShfTm(L, "Type") Then Thw CSub, "Give UdtStmt is invalid: No UdtnLn", "FstUdtStmt", StmtFst
O.Udtn = TakNm(L)
End Sub
Private Sub WSet_Mbr(O As TUdt, Stmty$()) ' Set O.Mbr() by the Middle part of Stmty (No Fst no las stmt)
Dim M() As TUmb, S$
Dim J%: For J = 1 To UB(Stmty) - 1
    S = BrkVmk(Stmty(J)).S1
    If S <> "" Then
        PushTUmb M, WTUbmStmt(S)
    End If
Next
O.Mbr = M
End Sub
Private Function WTUbmStmt(StmtUmb) As TUmb
Const CSub$ = CMod & "WTUbmStmt"
Dim L$: L = Brk1(StmtUmb, vbQuoSng).S1
Dim O As TUmb
O.Mbn = ShfNm(L)
O.IsAy = IsShfBkt(L)
L = LTrim(L)
If Not IsShfAs(L) Then Thw CSub, "Invalid StmtUmb, [ As ] is unexpected", "StmtUmb", StmtUmb
O.Tyn = L
WTUbmStmt = O
End Function
