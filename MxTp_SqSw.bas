Attribute VB_Name = "MxTp_SqSw"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_SqSw."
Private Type Pmsw: Pm As Dictionary: Sw As Dictionary: End Type 'Deriving(Ctor Ay)
Type SqSw: Pm As Dictionary: stmtSw As Dictionary: fldSw As Dictionary: End Type
Enum eSwOp: eOraSw: eEqnSw: End Enum
Function samp_sqtp_SqSw() As SqSw
With samp_sqtp_SqSw
Set .Pm = samp_sqtp_Pm
Set .fldSw = samp_sqtp_FldSw
Set .stmtSw = samp_sqtp_StmtSw
End With
End Function
Private Function samp_sqtp_StmtSw() As Dictionary

End Function
Private Function samp_sqtp_FldSw() As Dictionary

End Function
Function samp_sqtp_Pm() As Dictionary
Set samp_sqtp_Pm = DiLy(samp_sqtp_PmSrc)
End Function
Function samp_sqtp_PmSrc() As String()

End Function
Function samp_sqtp_SwSrc() As String()
Erase XX
X "?LvlY    EQ >>SumLvl Y"
X "?LvlM    EQ >>SumLvl M"
X "?LvlW    EQ >>SumLvl W"
X "?LvlD    EQ >>SumLvl D"
X "?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY"
X "?M       OR ?LvlD ?LvlW ?LvlM"
X "?W       OR ?LvlD ?LvlW"
X "?D       OR ?LvlD"
X "?Dte     OR ?LvlD"
X "?Mbr     OR >?BrkMbr"
X "?MbrCnt  OR >?BrkMbr"
X "?Div     OR >?BrkDiv"
X "?Sto     OR >?BrkSto"
X "?Crd     OR >?BrkCrd"
X "?#SEL#Div NE >LisDiv *blank"
X "?#SEL#Sto NE >LisSto *blank"
X "?#SEL#Crd NE >LisCrd *blank"
samp_sqtp_SwSrc = XX
Erase XX
End Function

Function sqSwLy(swLy$(), Pm As Dictionary) As SqSw
sqSwLy = UUEvl(WSwly(swLy), Pm)
End Function

Private Function UUEvl(L() As Swl, Pm As Dictionary) As SqSw
Const CSub$ = CMod & "UUEvl"
Dim OSw As New Dictionary
Dim J%
Do
    If Not UUEvlR(L, OSw, Pm) Then W2Thw L, OSw, Pm
    If SwlSI(L) = 0 Then UUEvl = W2Sw2(OSw): Exit Function
    ThwLoopTooMuch CSub, J
Loop
End Function
Private Function WSwly(swLy$()) As Swl()
Dim L: For Each L In Itr(swLy)
    PushSwl WSwly, WSwl(L)
Next
End Function
Private Function WSwl(Ln) As Swl
Const CSub$ = CMod & "WSwl"
Dim Ay$(): Ay = Tmy(Ln)
Dim Nm$: Nm = Ay(0)
Dim OpStr$: OpStr = Ay(1)
Dim Op As eBoolOp: Op = eBoolOp(OpStr)
Dim T1$, T2$
Select Case True
Case Op = eOpNe, Op = eOpEq:
    If Si(Ay) <> 4 Then Thw CSub, "Ln should have 4 terms for Eq | Ne", "Ln", Ln
    T1 = Ay(2): T2 = Ay(3):
Case Op = eOpAnd, Op = eOpOr
    If Si(Ay) < 3 Then Thw CSub, "Ln should have at 3 terms And | Or", "Ln", Ln
    Ay = AeFstN(Ay, 2)
End Select
WSwl = Swl(Nm, Op, Ay)
End Function

Private Sub W2Thw(L() As Swl, Sw As Dictionary, Pm As Dictionary)
Dim Left$(), Evld$()
    Left = FmtSwly(L)
    Evld = FmtDi(Sw)
Thw "Sw2zLy", "Cannot eval all Swl", "Some Swl is left", "SwLy Pm [Swl left] [evaluated Swl]", FmtSwly(L), Pm, Left, Evld
End Sub
Private Function W2Sw2(Sw As Dictionary) As SqSw
Dim Fld As New Dictionary
Dim StmtySrc As New Dictionary
Dim K: For Each K In Sw.Keys
    If HasPfx(K, "?") Then
        StmtySrc.Add RmvFst(K), Sw(K)
    Else
        Fld.Add K, Sw(K)
    End If
Next
Set W2Sw2.fldSw = Fld
Set W2Sw2.stmtSw = StmtySrc
End Function

Private Function UUEvlR(O() As Swl, OSw As Dictionary, Pm As Dictionary) As Boolean
Dim M As Swl, Ixy%(), OHasEvl As Boolean
Dim J%: For J = 0 To SwlUB(O)
    M = O(J)
    With UUEvlLn(M, Pmsw(Pm, OSw))
        If .Som Then
            OSw.Add M.Swn, .Bool
            OHasEvl = True
            PushI Ixy, J
        End If
    End With
Next
O = W3RmvSwl(O, Ixy)
UUEvlR = OHasEvl
End Function
Private Function W3RmvSwl(L() As Swl, Ixy%()) As Swl()
Dim J%: For J = 0 To SwlUB(L)
    If Not HasEle(Ixy, J) Then PushSwl W3RmvSwl, L(J)
Next
End Function
Private Function UUEvlLn(L As Swl, P As Pmsw) As Boolopt
With L
Select Case True
Case IsOrAndStr(.Op): UUEvlLn = UUEvlAndOr(CvOrAndStr(.Op), .Tml, P)
Case IsEqNeStr(.Op):  UUEvlLn = UUEvlEqNe(CvEqNeStr(.Op), .Tml(0), .Tml(1), P)
End Select
End With
End Function
Private Function UUEvlAndOr(Op As eOrAnd, Tml$(), P As Pmsw) As Boolopt
Dim O() As Boolean
Dim T: For Each T In Itr(Tml)
    With W5EvlTm(T, P)
        If Not .Som Then Exit Function
        PushI O, .Bool
    End With
Next
UUEvlAndOr = EvlBooly(O, Op)
End Function
Private Function W5EvlTm(T, P As Pmsw) As Boolopt
Select Case True
Case P.Sw.Exists(T): W5EvlTm = SomBool(P.Sw(T))
Case P.Pm.Exists(T): W5EvlTm = SomBool(P.Pm(T))
End Select
End Function

Private Function UUEvlEqNe(Op As eEqNe, T1$, T2$, P As Pmsw) As Boolopt 'Return True and set ORslt if evaluated
Dim S1$
    With W6EvlT1(T1, P.Pm)
        If Not .Som Then Exit Function
        S1 = .Str
    End With
Dim S2$
    With W6EvlT2(T2, P.Pm)
        If Not .Som Then Exit Function
        S2 = .Str
    End With
Select Case True
Case Op = eEqnEq: UUEvlEqNe = SomBool(S1 = S2)
Case Op = eEqnNe: UUEvlEqNe = SomBool(S1 <> S2)
Case Else: ThwImposs CSub
End Select
End Function
Private Function W6EvlT1(T1$, Pm As Dictionary) As Stropt
If Pm.Exists(T1) Then W6EvlT1 = SomStr(Pm(T1))
End Function
Private Function W6EvlT2(T2$, Pm As Dictionary) As Stropt
If T2 = "*Blank" Then W6EvlT2 = SomStr(""): Exit Function
Dim M As Stropt: M = W6EvlT1(T2, Pm)
If M.Som Then W6EvlT2 = M: Exit Function
W6EvlT2 = SomStr(T2)
End Function

Private Function Pmsw(Pm As Dictionary, Sw As Dictionary) As Pmsw
With Pmsw
    Set .Pm = Pm
    Set .Sw = Sw
End With
End Function
Function PmswAdd(A As Pmsw, B As Pmsw) As Pmsw(): PushPmsw PmswAdd, A: PushPmsw PmswAdd, B: End Function
Sub PushPmswAy(O() As Pmsw, A() As Pmsw): Dim J&: For J = 0 To PmswUB(A): PushPmsw O, A(J): Next: End Sub
Sub PushPmsw(O() As Pmsw, M As Pmsw): Dim N&: N = PmswSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function PmswSI&(A() As Pmsw): On Error Resume Next: PmswSI = UBound(A) + 1: End Function
Function PmswUB&(A() As Pmsw): PmswUB = PmswSI(A) - 1: End Function
