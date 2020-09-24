Attribute VB_Name = "MxDao_Sql_FldLis"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_FldLis."
Private Enum eFldLisTy: eFldLisTySel: eFldLisTySet: eFldLisTySelDis: End Enum
Private Const EnmqssFldLisTy$ = "eFldLisTy? Sel Set SelDis"
Private Function FmtFldLis$(FldLis$, T As eFldLisTy)
'Always return vbcrLf and 7-spc each lines, and no space in front
Dim Brk$
    Select Case T
    Case eFldLisTySel, eFldLisTySelDis: Brk = " As "
    Case eFldLisTySet: Brk = "="
    Case Else: ThwEnm CSub, T, EnmqssFldLisTy
    End Select
Dim S1$()
Dim S2$()
    Dim Fldy$(): Fldy = SplitFldLis(FldLis)
    Dim S() As S12: S = S12ySyBrk2(Fldy, Brk)
    S1 = AmAli(S1y(S))
    S2 = AmAli(S2y(S))
Dim OLy$()
    If Brk = "=" Then Brk = " = "
    Dim U%: U = UB(S1)
    ReDim OLy(U)
    Dim P$: P = Space(Len(Brk))
    Dim J%: For J = 0 To U
        If Trim(S1(J)) = "" Then
            OLy(J) = S1(J) & P & S2(J)   '<==
        Else
            OLy(J) = S1(J) & Brk & S2(J) '<===
        End If
    Next

'FmtFldLis:
    Dim Sep$
    Select Case T
    Case eFldLisTySel:                     Sep = vbCmaCrLf & Space(7)
    Case eFldLisTySelDis, eFldLisTySet: Sep = vbCmaCrLf & vbSpc4
    End Select
    If T = eFldLisTySelDis Then
        P = vbCrLf & "    "
    Else
        P = ""
    End If
FmtFldLis = P & Jn(OLy, Sep) '<===
'Debug.Print FmtFldLis: Stop
End Function

Function SplitFldLis(FldLis$) As String()
Dim L$: L = FldLis
While LTrim(L) <> ""
    PushI SplitFldLis, ShfFldLisItm(L)
Wend
End Function
Private Function ShfFldLisItm$(OLn$)
Dim L$: L = LTrim(OLn)
Dim P%: P = 1
Dim W%: W = Len(OLn)
While P <= W
    Dim J%: J = J + 1: If J > 500 Then RaiseMsg CSub & ": looping too much"
    Dim C$: C = Mid(OLn, P, 1)
    Select Case C
    Case "("
        P = PosBktCls(L, P): If P = 0 Then Thw CSub, "Inbalance bracket in @OLn", "@OLn", OLn
        P = P + 1
    Case vbQuoSng, vbQuoDbl
        P = InStr(P + 1, L, C): If P = 0 Then Thw CSub, "Inbalance Quo in @OLn", "@OLn Quo", OLn, C
        P = P + 1
    Case ","
        ShfFldLisItm = RTrim(Left(L, P - 1))
        OLn = LTrim(Mid(L, P + 1))
        Exit Function '<======================
    Case Else
        P = P + 1
    End Select
Wend
ShfFldLisItm = RTrim(L)
OLn = ""
End Function
Function FmtFldLisSet$(FldLisoSet$):          FmtFldLisSet = FmtFldLis(FldLisoSet, eFldLisTySet):       End Function
Function FmtFldLisSel$(FldLisoSel$):          FmtFldLisSel = FmtFldLis(FldLisoSel, eFldLisTySel):       End Function
Function FmtFldLisSelDis$(FldLisoSelDis$): FmtFldLisSelDis = FmtFldLis(FldLisoSelDis, eFldLisTySelDis): End Function
