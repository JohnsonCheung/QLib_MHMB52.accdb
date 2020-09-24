Attribute VB_Name = "MxVb_Str_Shf"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Shf."
Function ShfDotSeg$(OLn$)
ShfDotSeg = ShfBef(OLn, ".")
End Function

Function ShfEq(OLn$) As Boolean: ShfEq = IsShfTm(OLn, "="): End Function

Private Sub B_ShfBetBkt()
Dim A$, Ept1$
A$ = "(O$()) As X": Ept = "O$()": Ept1 = " As X": GoSub Tst
Exit Sub
Tst:
    Act = ShfBetBkt(A)
    C
    Ass A = Ept1
    Return
End Sub
Function ShfBetBkt$(OLn$, Optional BktOpn$ = vbBktOpn)
If ChrFst(OLn) <> BktOpn Then Exit Function
ShfBetBkt = BetBkt(OLn, BktOpn)
If ShfBetBkt <> "" Then OLn = AftBkt(OLn, BktOpn)
End Function

Private Sub B_ShfBef()
Dim L$, Sep$, EptL$, InlSep As Boolean, Cas As eCas
'GoSub T1
GoSub T2
'GoSub T3
Exit Sub
T1:
    L = " aaa == bbb"
    Sep = "=="
    InlSep = True
    Ept = " aaa =="
    EptL = " bbb"
    GoTo Tst
T2:
    L = " aaa == bbb"
    Sep = "=="
    InlSep = False
    Ept = " aaa "
    EptL = "== bbb"
    GoTo Tst
T3:
    L = "aaa.bbb"
    Sep = "."
    Ept = "aaa"
    EptL = ".bbb"
    GoTo Tst
Tst:
    Act = ShfBef(L, Sep, InlSep, Cas)
    C
    If L <> EptL Then Stop
    Return
End Sub
Function ShfBefEq$(OLn$, Optional InlSep As Boolean): ShfBefEq = ShfBef(OLn, "=", InlSep): End Function
Function ShfPfx$(OLn$, Pfx$, Optional C As eCas)
ShfPfx = TakPfx(OLn, Pfx, C): If ShfPfx <> "" Then OLn = Mid(OLn, Len(Pfx) + 1)
End Function
Function ShfPfxy$(OLn$, Pfxy$())
ShfPfxy = PfxPfxy(OLn, Pfxy): If ShfPfxy = "" Then Exit Function
OLn = RmvPfx(OLn, ShfPfxy)
End Function
Function ShfPfxSpc$(OLn$, Pfx, Optional C As eCas): ShfPfxSpc = ShfPfx(OLn, Pfx & " ", C): End Function
Function ShfPfxySpc$(OLn$, Pfxy$(), Optional C As eCas = eCasIgn)
ShfPfxySpc = PfxPfxySpc(OLn, Pfxy, C): If ShfPfxySpc = "" Then Exit Function
OLn = RmvPfxSpc(OLn, ShfPfxySpc)
End Function
Function ShfBef$(OLn$, Bef, Optional InlSep As Boolean, Optional C As eCas)
Const CSub$ = CMod & "ShfBef"
Dim P&: P = InStr(1, OLn, Bef, VbCprMth(C)): If P = 0 Then Thw CSub, "Bef not found in OLn", "Bef OLn", Bef, OLn
If InlSep Then
    Dim L%: L = Len(Bef)
    ShfBef = Left(OLn, P - 1 + L)
    OLn = Mid(OLn, P + L)
Else
    ShfBef = Left(OLn, P - 1)
    OLn = Mid(OLn, P)
End If
End Function
Function ShfBefIf$(OLn$, Bef, Optional InlSep As Boolean)
Dim P&: P = InStr(OLn, Bef): If P = 0 Then Exit Function
ShfBefIf = ShfBef(OLn, Bef, InlSep)
End Function
Function ShfBefOrAll$(OLn$, Sep$, Optional NoTrim As Boolean)
Dim P%: P = InStr(OLn, Sep)
If P = 0 Then
    If NoTrim Then
        ShfBefOrAll = OLn
    Else
        ShfBefOrAll = Trim(OLn)
    End If
    OLn = ""
    Exit Function
End If
ShfBefOrAll = Bef(OLn, Sep, NoTrim)
OLn = Aft(OLn, Sep, NoTrim)
End Function
Function ShfTmBefDot$(OLn$)
With Brk2Dot(OLn, NoTrim:=True)
    ShfTmBefDot = .S1
    OLn = .S2
End With
End Function
