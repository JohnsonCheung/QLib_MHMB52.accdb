Attribute VB_Name = "MxDta_Da_Dy_Op_DyAddDc"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Op_AddDc."
Function DyAddDcStr12(SCol1$(), SCol2$()) As String()
Dim A$(): A = AmAli(SCol1)
Dim J&: For J = 0 To UB(A)
    PushI DyAddDcStr12, A(J) & SCol2(J)
Next
End Function
Function DyAddDcIx(Dy(), NoIx As Boolean, Optional NHdr%, Optional ChrUL$ = "=") As Variant()
If NoIx Then DyAddDcIx = Dy: Exit Function
Dim DcIx1: DcIx1 = DcIx(Si(Dy), NHdr, ChrUL)
For J = 0 To UB(Dy)
    PushI DyAddDcIx, AyItmAy(DcIx(J), Dy(J))
Next
End Function
Private Function DcIx(NDy&, NHdr%, ChrUL$) As String()
Dim HdrHsh$: HdrHsh = String(NDig(NDy), "#")
Select Case NHdr
Case 0: DcIx = IntySno(NDy, 1)
Case 1: DcIx = SyStrSy(HdrHsh, IntySno(NDy, 1))
Case Is > 1
    Dim HdryEmp$()
    DcIx = SyAp(HdryEmp, HdrHsh, IntySno(NDy, 1))
Case Else: ThwPm "NHdr must>=0", "NHdr", NHdr
End Select
End Function
Function DyAddDc(Dy(), ValToBeAddAsLasCol) As Variant()
'Ret : a new :Dy with a col of value all eq to @ValToBeAddAsLasCol at end
Dim O(): O = Dy
Dim ToU&
    ToU = NDcDy(Dy)
Dim J&, Dr
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    Dr(ToU) = ValToBeAddAsLasCol
    O(J) = Dr
    J = J + 1
Next
DyAddDc = O
End Function

Function DyAddDcN(Dy(), Optional ByNCol% = 1) As Variant()
Dim NewU&
    NewU = NDcDy(Dy) + ByNCol - 1
Dim O()
    Dim UDy&: UDy = UB(Dy)
    O = AyReDim(O, UDy)
    Dim J&
    For J = 0 To UDy
        O(J) = AyReDim(Dy(J), NewU)
    Next
DyAddDcN = O
End Function

Function DyAddDcC(Dy(), C) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim O(): O = DyAddDcN(Dy)
    Dim UCol%: UCol = UB(Dy(0))
    Dim J&
    For J = 0 To UB(Dy)
       O(J)(UCol) = C
    Next
DyAddDcC = O
End Function

Function DyAddDcMap(A As Drs, NewFeqFunQuoFmFldSsl$) As Drs
Dim NewFldVy(), FmVy()
Dim I, S$, NewFld$, Fun$, FmFld$
For Each I In SySs(NewFeqFunQuoFmFldSsl)
    S = I
    NewFld = Bef(S, "=")
    Fun = IsBet(S, "=", "(")
    FmFld = BetBkt(S)
    FmVy = DcDrs(A, FmFld)
    NewFldVy = AyMap(FmVy, Fun)
    Stop '
Next
End Function


Function RixWhColEq&(Dy, Cix, Eqval)
Dim R&
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(Cix) = Eqval Then RixWhColEq = R
    R = R + 1
Next
RixWhColEq = -1
End Function
Function HasColEq(A As Drs, C$, V) As Boolean
HasColEq = HasColEqDy(A.Dy, IxEle(A.Fny, C), V)
End Function

Function RmvPfxDrs(A As Drs, C$, Pfx$) As Drs
Dim Dr, ODy(), J&, I%
ODy = A.Dy
I = IxEle(A.Fny, C)
For Each Dr In Itr(A.Dy)
    Dr(I) = RmvPfx(Dr(I), Pfx)
    ODy(J) = Dr
    J = J + 1
Next
RmvPfxDrs = Drs(A.Fny, ODy)
End Function

Function RxyeDyVy(Dy(), Vy) As Long()
'Fm Dy: ! to be selected if it ne to @Vy.  It has sam NCol as Si-Vy
'Fm Vy : ! to select @Dy
'Ret   : Rxy of @Dy if the rec ne @Vy
Dim Rix&, Dr: For Each Dr In Itr(Dy)
    If Not IsEqAy(Dr, Vy) Then PushI RxyeDyVy, Rix
    Rix = Rix + 1
Next
End Function

Function RxywDyVy(Dy(), Vy) As Long()
'Fm Dy: ! to be selected if it eq to @Vy.  It has sam NCol as Si-Vy
'Fm Vy : ! to select @Dy
'Ret   : Rxy of @Dy if the rec eq @Vy
Dim Rix&, Dr: For Each Dr In Itr(Dy)
    If IsEqAy(Dr, Vy) Then PushI RxywDyVy, Rix
    Rix = Rix + 1
Next
End Function

Function DwTopN(A As Drs, Optional N = 50) As Drs
If N <= 0 Then DwTopN = A: Exit Function
DwTopN = Drs(A.Fny, CvAv(AwFstN(A.Dy, N)))
End Function

Function ValDrs(A As Drs, C$, V, Coln$)
Const CSub$ = CMod & "ValDrs"
Dim Dr, Ix%, IxRet%
Ix = IxEle(A.Fny, C)
IxRet = IxEle(A.Fny, Coln)
For Each Dr In Itr(A.Dy)
    If Dr(Ix) = V Then
        ValDrs = Dr(IxRet)
        Exit Function
    End If
Next
Thw CSub, "In Drs, there is no record with DcDrs-A eq Value-B, so no DcDrs-C is returened", "DcDrs-A Value-B DcDrs-C Drs-Fny Drs-NRec", C, V, Coln, A.Fny, NRecDrs(A)
End Function

Function DyAddDc2(Dy(), V1, V2) As Variant():     DyAddDc2 = DyAddDcDr(Dy, Av(V1, V2)):     End Function
Function DyAddDc3(Dy(), V1, V2, V3) As Variant(): DyAddDc3 = DyAddDcDr(Dy, Av(V1, V2, V3)): End Function


Function DyAddDcDr(Dy(), DrToAdd(), Optional IxBef = 0) As Variant()
Const CSub$ = CMod & "DyAddDcFldVy"
Dim Dr, J&, O(), U&
U = UB(DrToAdd)
If U = -1 Then Exit Function
If U <> UB(Dy) Then Thw CSub, "Row-in-Dy <> Si-Fvy", "Row-in-Dy Si-Fvy", Si(Dy), Si(DrToAdd)
ReDim O(U)
Dim UNw%: UNw = UB(Dy(0)) + Si(DrToAdd)
For Each Dr In Itr(Dy)
    If Si(Dr) > IxBef Then Thw CSub, "Some Dr in Dy has bigger size than AtIx", "DRs AtIx", Si(Dr), IxBef
    ReDim Preserve Dr(UNw)
    Dr(IxBef) = DrToAdd(J)
    O(J) = DrToAdd
    J = J + 1
Next
DyAddDcDr = O
End Function
