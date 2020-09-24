Attribute VB_Name = "MxDta_Da_Op_DyOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dy."
Function DyReDimDc(Dy()) As Variant()
Dim UDc%: UDc = UDcDy(Dy)

Dim Dr: For Each Dr In Itr(Dy)
    PushI DyReDimDc, AyReDim(Dr, UDc)
Next
End Function
Function IsCnstCol(Dy(), C%) As Boolean
Const CSub$ = CMod & "IsCnstCol"
If Si(Dy) = 0 Then Thw CSub, "No record in Dy"
Dim V: V = Dy(0)(C)
Dim Dr: For Each Dr In Dy
    If UB(Dr) >= C Then
        If V <> Dr(C) Then Exit Function
    Else
        If V <> Empty Then Exit Function
    End If
Next
IsCnstCol = True
End Function
Function JnDotDy(Dy()) As String()
Dim Dr: For Each Dr In Itr(Dy)
    PushI JnDotDy, JnDot(Dr)
Next
End Function

Function DyAliR(Dy(), C) As Variant()
Dim Ay$(): Ay = AmAliR(DcDy(Dy, C))
Dim O(): O = Dy
Dim J&
For J = 0 To UB(O)
    O(J)(C) = Ay(J)
Next
DyAliR = O
End Function

Function DicntDyC(Dy(), C&) As Dictionary
Set DicntDyC = DiCnt(DcDy(Dy, C))
End Function

Function DcyDy(Dy()) As Variant()
Dim J%: For J = 0 To NDcDy(Dy)
    PushI DcyDy, DcDy(Dy, J)
Next
End Function

Function DisCol(A As Drs, C$)
DisCol = AwDis(DcDy(A.Dy, IxEle(A.Fny, C)))
End Function

Function DisSCol(A As Drs, C$) As String()
Dim I%: I = IxEle(A.Fny, C)
Dim DcDrs$(): DcDrs = DcStrDy(A.Dy, I)
DisSCol = AwDis(DcDrs)
End Function

Function DistDcDy(Dy(), C&) As Variant()
DistDcDy = AwDis(DcDy(Dy, C))
End Function

Function DotSyDy(Dy()) As String()
Dim Dr
For Each Dr In Itr(Dy)
    PushI DotSyDy, JnDot(Dr)
Next
End Function

Function DyJnFldNFld(Dy(), FstNFld%, Optional Sep$ = " ") As Variant()
Dim U%: U = FstNFld - 1
Dim UK%: UK = U - 1
Dim O(), Dr
For Each Dr In Itr(Dy)
    If U <> UB(Dr) Then
        ReDim Preserve Dr(U)
    End If
    Dim Ix: Ix = RixDyDr(O, AwFstN(Dr, UK))
    If Ix = -1 Then
        PushI O, Dr
    Else
        Stop
'        O(Ix)(U) = AddNB(O(Ix)(U), Sep) & Dr(U)
    End If
Next
DyJnFldNFld = O
End Function

Function DyJnFldKK(Dy(), KKIxy&(), JnFldIx&) As Variant()
'Ret : :Dy-@KKIxy-@JnFldIx ! Ret Dy of Si(@KKIxy) + 1 columns with UKey-KKIxy
Dim Ixy&(): Ixy = KKIxy: PushI Ixy, JnFldIx
Dim N%: N = Si(Ixy)
DyJnFldKK = DyJnFldNFld(SelCol(Dy, Ixy), N)
End Function

Function DySq(Sq()) As Variant()
If Si(Sq) = 0 Then Exit Function
Dim R&: For R = 1 To UBound(Sq, 1)
    PushI DySq, DrSq(Sq, R)
Next
End Function

Function DySsy(Ssy$()) As Variant()
Dim Ss: For Each Ss In Itr(Ssy)
    PushI DySsy, SySs(Ss)
Next
End Function

Function DyWhRxayAnd(Dy(), C%, RxayAnd() As RegExp) As Variant()
If Si(RxayAnd) = 0 Then DyWhRxayAnd = Dy: Exit Function
Dim Dr: For Each Dr In Dy
    If UB(Dr) >= C Then
        If HasRxAyAnd(Dr(C), RxayAnd) Then PushI DyWhRxayAnd, Dr
    End If
Next
End Function
Function DyWh2Ne(Dy(), C1&, C2&) As Variant()
Dim Dr: For Each Dr In Dy
    If Dr(C1) <> Dr(C2) Then PushI DyWh2Ne, Dr
Next
End Function

Function DyWhGt(Dy(), C%, GtV) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) > GtV Then PushI DyWhGt, Dr
Next
End Function

Function DyWhCoLiny(Dy(), Cix%, InAy) As Variant()
Const CSub$ = CMod & "DyWhCoLiny"
If Not IsArray(InAy) Then Thw CSub, "[InAy] is not Array, but [TypeName]", "InAy-TypeName", TypeName(InAy)
If Si(InAy) = 0 Then DyWhCoLiny = Dy: Exit Function
Dim Dr
For Each Dr In Itr(Dy)
    If HasEle(InAy, Dr(Cix)) Then PushI DyWhCoLiny, Dr
Next
End Function

Function DyWhNe(Dy(), C, Ne) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(C) <> Ne Then PushI DyWhNe, Dr
Next
End Function

Function DyWhDis(Dy()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushNoDupDr DyWhDis, Dr
Next
End Function

Function DyWhDup1(Dy(), C&) As Variant()
Dim Dup$(), Dr, O()
Dup = CvSy(AwDup(DcStrDy(Dy, C)))
For Each Dr In Itr(Dy)
    If HasEle(Dup, Dr(C)) Then PushI DyWhDup1, Dr
Next
End Function


Function DySqCnoy(Sq(), Cnoy%()) As Variant()
'Fm Cnoy : :Cnoy ! selecting which col of @Sq
Dim R&: For R = 1 To UBound(Sq, 1)
    PushI DySqCnoy, DrSqCnoy(Sq, R, Cnoy)
Next
End Function


Function DyVbl(Vbl$) As Variant()
Dim L: For Each L In Itr(SplitVBar(Vbl))
    PushI DyVbl, SySs(L)
Next
End Function

Function DrDyFstEq(Dy(), C, V) As Variant()
Const CSub$ = CMod & "DrDyFstEq"
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(C) = V Then DrDyFstEq = Dr: Exit Function
Next
Thw CSub, "No first Dr in Dy of Cix eq to V", "Cix V Dy", C, V, LyDy(Dy)
End Function

Function FstRecEqDy(Dy(), C, Eq, SelIxy&()) As Variant()
Const CSub$ = CMod & "FstRecEqDy"
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) = Eq Then FstRecEqDy = Array(AwIxy(Dr, SelIxy)): Exit Function
Next
Thw CSub, "No first rec in Dy of DcDrs-A eq to Val-B", "DcDrs-A Val-B Dy", C, Eq, LyDy(Dy)
End Function

Function HasColEqDy(Dy(), C&, Eq) As Boolean
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) = Eq Then HasColEqDy = True: Exit Function
Next
End Function

Function HasDr(Dy(), Dr) As Boolean
Dim IDr
For Each IDr In Itr(Dy)
    If IsEqAy(IDr, Dr) Then HasDr = True: Exit Function
Next
End Function

Function HasDrIxy(Dy(), Dr, Ixy&()) As Boolean
Dim IDr, A()
For Each IDr In Itr(Dy)
    A = AwIxy(IDr, Ixy)
    If IsEqAy(A, IDr) Then HasDrIxy = True: Exit Function
Next
End Function

Function IsRowBrk(Dy(), R&, BrkColIx&) As Boolean
If Si(Dy) = 0 Then Exit Function
If R& = 0 Then Exit Function
If R = UB(Dy) Then Exit Function
If Dy(R)(BrkColIx) = Dy(R - 1)(BrkColIx) Then Exit Function
IsRowBrk = True
End Function

Function IxyCnoy(Cnoy) As Long()
If Si(Cnoy) = 0 Then Exit Function
Dim Cno: For Each Cno In Cnoy
    PushI IxyCnoy, Cno - 1
Next
End Function

Function DyKeepFstNDc(Dy(), N%) As Variant()
Dim Dr, U%
U = N - 1
For Each Dr In Itr(Dy)
    ReDim Preserve Dr(U)
    PushI DyKeepFstNDc, Dr
Next
End Function

Function DrsKeepFstNDc(A As Drs, N%) As Drs
DrsKeepFstNDc = Drs(CvSy(AwFstN(A.Fny, N)), DyKeepFstNDc(A.Dy, N))
End Function

Function UDcDy%(Dy()): UDcDy = NDcDy(Dy) - 1: End Function
Function NDcDy%(Dy())
Dim O%, Dr: For Each Dr In Itr(Dy)
    O = Max(O, Si(Dr))
Next
NDcDy = O
End Function

Function RxyDyDr(Dy(), Dr) As Long() 'ret :Rxy of @Dy for those *IDr = @Dr
Dim NCol%: NCol = Si(Dr)
Dim IDr, Rix&: For Each IDr In Itr(Dy)
    If IsEqAy(AwFstN(IDr, NCol), Dr) Then
        PushI RxyDyDr, Rix
    End If
    Rix = Rix + 1
Next
End Function

Function RixDyDr&(Dy(), Dr) 'ret :Rix of @Dy for its fst *IDr = @Dr
Dim NCol%: NCol = Si(Dr)
Dim IDr, Rix&: For Each IDr In Itr(Dy)
    If IsEqAy(AwFstN(IDr, NCol), Dr) Then
        RixDyDr = Rix
        Exit Function
    End If
    Rix = Rix + 1
Next
RixDyDr = -1
End Function

Sub ChkIsEqDy(Dy(), B(), Optional Fun$)
If Not IsEqDy(Dy, B) Then Thw Fun, "2 Dy not Eq", "Dy1 Dy2", FmtDy(Dy), FmtDy(B)
End Sub

Private Sub B_DyJnFldKK()
GoSub T1
Dim Dy(), KKIxy&(), JnFldIx&, Sep$
Exit Sub
T1:
    Dy = Array(Array(1, 2, 3, 4, "Dy"), Array(1, 2, 3, 6, "B"), Array(1, 2, 2, 8, "C"), Array(1, 2, 2, 12, "DD"), _
    Array(2, 3, 1, 1, "x"), Array(12, 3), Array(12, 3, 1, 2, "XX"))
    Ept = Array()
    KKIxy = Array(0, 1, 2)
    JnFldIx = 4
    GoTo Tst
Tst:
    Act = DyJnFldKK(Dy, KKIxy, JnFldIx)
    BrwDy CvAv(Act)
    StopNE
    Return
End Sub

Function IsDrsEqCol(D As Drs, C1$, C2$) As Boolean ' Is col-@C1 equal col-@C2
Dim A%: A = CixDrs(D, C1)
Dim B%: B = CixDrs(D, C2)
IsDrsEqCol = IsDyEqCol(D.Dy, A, B)
End Function
Function IsDyEqCol(Dy(), C1%, C2%) As Boolean ' Is col-@C1 equal col-@C2
Dim M%: M = Max(C1, C2)
Dim Dr: For Each Dr In Itr(Dy)
    If UB(Dr) < M Then Exit Function  ' Dr may have less field than C1,C2
    If Dr(C1) <> Dr(C2) Then Exit Function
Next
IsDyEqCol = True
End Function
