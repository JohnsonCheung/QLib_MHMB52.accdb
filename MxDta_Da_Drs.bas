Attribute VB_Name = "MxDta_Da_Drs"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Drs."
Private Type GpDrs
    GpDrs As Drs
    RLvlGpIx() As Long
End Type

Function DrsDt(A As Dt) As Drs: DrsDt = Drs(A.Fny, A.Dy): End Function
Function DrsPutDc(D As Drs, C$, Dc) As Drs
Const CSub$ = CMod & "DrsPutDc"
Dim Dy(): Dy = D.Dy
If Si(Dy) <> Si(Dc) Then Thw CSub, "@Drs.Dy and @Dc Should be same size", "@Drs-Dy-Si @Dc-Si", Si(Dy), Si(Dc)
Dim DyNw()
    Dim Dr, J&: For Each Dr In Itr(Dy)
        PushI Dr, Dc(J)
        PushI DyNw, Dr
        J = J + 1
    Next
DrsPutDc = DrsFfAdd(D, C, DyNw)
End Function

Sub AsgTJn(TmlJn$, OFnyJnA$(), OFnyJnB$())
Dim SyJn$(): Stop 'SyJn = Tml(TmlJn)
Dim U%: U = UB(SyJn)
ReDim OFnyJnA(U)
ReDim OFnyJnB(U)
Dim J: For J = 0 To U
    With BrkBoth(SyJn(J), ":")
        OFnyJnA(J) = .S1
        OFnyJnB(J) = .S2
    End With
Next
End Sub

Function DrsFfAdd(D As Drs, FfAdd$, DyNw()) As Drs: DrsFfAdd = Drs(SyAddSS(D.Fny, FfAdd), DyNw): End Function

Function DrsAdd(A As Drs, B As Drs) As Drs
Const CSub$ = CMod & "DrsAdd"
If IsEmpDrs(A) Then DrsAdd = B: Exit Function
If IsEmpDrs(B) Then DrsAdd = A: Exit Function
If Not IsEqAy(A.Fny, B.Fny) Then Thw CSub, "Dif Fny: Cannot add", "A-Fny B-Fny", A.Fny, B.Fny
DrsAdd = Drs(A.Fny, AvAdd(A.Dy, B.Dy))
End Function

Function DrsAdd3(A As Drs, B As Drs, C As Drs) As Drs
Const CSub$ = CMod & "DrsAdd3"
ChkIsEqAy A.Fny, B.Fny, "FnyA FnyB", CSub
ChkIsEqAy B.Fny, C.Fny, "FnyB FnyC", CSub
DrsAdd3 = A
AppDy DrsAdd3.Dy, B.Dy
AppDy DrsAdd3.Dy, C.Dy
End Function

Sub AppDrs(O As Drs, M As Drs)
Const CSub$ = CMod & "AppDrs"
ChkIsEqAy O.Fny, M.Fny, "FnyO FnyM", CSub
ChkIsEqAy O.Fny, M.Fny, "FnyO FnyM", CSub
AppDy O.Dy, M.Dy
End Sub
Sub AppDy(O(), M())
Dim UO&, UM&, U&, J&
UO = UB(O)
UM = UB(M)
U = UO + UM + 1
ReDim Preserve O(U)
For J = UO + 1 To U
    O(J) = M(J - UO - 1)
Next
End Sub
Sub AppDrsSub(O As Drs, M As Drs)
Dim Ixy&(): Ixy = IxyEley(O.Fny, M.Fny)
Dim ODy(): ODy = O.Dy
Dim Dr
For Each Dr In Itr(M.Dy)
    PushI ODy, SelDr(CvAv(Dr), Ixy)
Next
O.Dy = ODy
End Sub

Sub AsgCol(A As Drs, CC$, ParamArray OColAp())
Dim OColAv(), J%, DcDrs, C$()
OColAv = OColAp
C = SySs(CC)
For J = 0 To UB(OColAv)
    DcDrs = IntoDrsC(OColAv(J), A, C(J))
    OColAp(J) = DcDrs
Next
End Sub

Sub AsgColDist(A As Drs, CC$, ParamArray OColAp())
Dim OColAv(), J%, DcDrs, B As Drs, C$()
B = DwDist(A, CC)
OColAv = OColAp
C = SySs(CC)
For J = 0 To UB(OColAv)
    DcDrs = IntoDrsC(OColAv(J), B, C(J))
    OColAp(J) = DcDrs
Next
End Sub

Function AvDrsC(A As Drs, C) As Variant(): AvDrsC = IntoDrsC(Array(), A, C): End Function
Function CntLyCntDi(DiCnt As Dictionary, CntWdt%) As String()
Dim K
For Each K In DiCnt.Keys
    PushI CntLyCntDi, AliR(DiCnt(K), CntWdt) & " " & K
Next
End Function

Sub ColApDrs(A As Drs, CC$, ParamArray OColAp())
Dim Av(): Av = OColAp
Dim C$(): C = SySs(CC)
Dim J%, O
For J = 0 To UB(Av)
    O = OColAp(J)
    O = IntoDrsC(O, A, C(J)) 'Must put into O first!!
                              'This will die: OColAp(J) = IntoDrsC(O, A, C(J))
    OColAp(J) = O
Next
End Sub

Function ColGp(DcDrs(), RLvlGpIx&()) As Variant()
'Fm DcDrs      : DcDrs to gp
'Fm RLvlGpIx : Each V in DcDrs is mapped to GpIx by this RLvlGpix @@
Const CSub$ = CMod & "ColGp"
ChkIsEqAySi DcDrs, RLvlGpIx, CSub
Dim MaxGpIx&: MaxGpIx = EleMax(RLvlGpIx)
Dim O(): ReDim O(MaxGpIx)
Dim I&: For I = 0 To MaxGpIx
    O(I) = Array()
Next
I = 0
Dim V: For Each V In Itr(DcDrs)
    Dim GpIx&: GpIx = RLvlGpIx(I)
    PushI O(GpIx), V
    I = I + 1
Next
ColGp = O
End Function

Function DicItmWdt%(A As Dictionary)
Dim I, O%
For Each I In A.Items
    O = Max(Len(I), O)
Next
DicItmWdt = O
End Function

Function DiRenFf(RenFf$) As Dictionary
Const CSub$ = CMod & "DiRenFf"
Set DiRenFf = New Dictionary
Dim Ay$(): Ay = SySs(RenFf)
Dim V: For Each V In SySs(RenFf)
    If HasSsub(V, ":") Then
        DiRenFf.Add Bef(V, ":"), Aft(V, ":")
    Else
        Thw CSub, "Invalid RenFf.  all Sterm has have [:]", "RenFf", RenFf
    End If
Next
End Function

Function DrsInsDcAft(A As Drs, C$, V, FldnAft$) As Drs: DrsInsDcAft = X_DrsInsDc(A, C, V, True, FldnAft):  End Function
Function DrsInsDcBef(A As Drs, C$, V, FldnBef$) As Drs: DrsInsDcBef = X_DrsInsDc(A, C, V, False, FldnBef): End Function
Private Function X_DrsInsDc(D As Drs, C$, V, IsAft As Boolean, Fldn$) As Drs
Dim Fny$(), Dy(), Ix%, Fny1$()
Fny = D.Fny
Ix = IxEle(Fny, C): If Ix = -1 Then Stop
If IsAft Then
    Ix = Ix + 1
End If
Dim FnyNw$(): FnyNw = AyIns(Fny, Fldn, CLng(Ix))
Dy = DyInsDc(D.Dy, V, Ix)
X_DrsInsDc = Drs(FnyNw, Dy)
End Function

Function FfDrs$(D As Drs): FfDrs = Join(D.Fny): End Function


Function DrsFillDcOfEmpCellbyAbove(D As Drs, C$) As Drs
'Fm D : It has a str col C
'Ret  : Fill in the blank col-C val by las val
Dim Cix%: Cix = IxEle(D.Fny, C)
DrsFillDcOfEmpCellbyAbove = Drs(D.Fny, WDy(D.Dy, Cix))
End Function
Private Sub B_WDy()
'GoSub Try
GoSub T1
GoSub T2
Exit Sub
Dim Dy(), Cix%
T1:
    Dy = Array(Array(1, 2), Array(1, 2, Empty), Array(1, 2, Empty), Array(1, 2, 4))
    Cix = 2
    Ept = Array(Array(1, 2, Empty), Array(1, 2, Empty), Array(1, 2, Empty), Array(1, 2, 4))
    GoTo Tst
T2:
    Dy = Array(Array(1, 2, 3), Array(1, 2, Empty), Array(1, 2, Empty), Array(1, 2, 4))
    Cix = 2
    Ept = Array(Array(1, 2, 3), Array(1, 2, 3), Array(1, 2, 3), Array(1, 2, 4))
    GoTo Tst
Tst:
    Act = WDy(Dy, Cix)
    C
    Return
Try:
    Dy = Array(Array(1, 2), Array(1, 2, Empty), Array(1, 2, Empty), Array(1, 2, 4))
    Stop
    MsgBox Dy(1)(2)
    Dy(1)(2) = 1
    Stop
    MsgBox Dy(1)(2)
    Return
   
End Sub
Private Function WDy(Dy(), Cix%) As Variant()
If IsEmpAy(Dy) Then Exit Function
Dim O(): O = Dy
Dim VLas: If UB(Dy(0)) >= Cix Then VLas = Dy(0)(Cix)
Dim Dr, VCur: For Each Dr In Itr(Dy)
    If UB(Dr) < Cix Then ReDim Preserve Dr(Cix)
    VCur = Dr(Cix)
    If IsEmpty(VCur) Then
        Dr(Cix) = VLas
        O(Cix) = Dr
    Else
        VLas = VCur
    End If
Next
WDy = O
End Function

Function DrsMapAy(Ay, MapFunNN$, Optional FF$, Optional ValNm$ = "V") As Drs: DrsMapAy = DrsMapItr(Itr(Ay), MapFunNN, FF, ValNm): End Function
Function DrsMapItr(Itr, MapFunNN$, Optional Ff0$, Optional ValNm$ = "V") As Drs
Dim Dy(), V: For Each V In Itr
    Dim Dr(): Dr = Array(V)
    Dim F: For Each F In ItrSS(MapFunNN)
        PushI Dr, Run(F, V)
    Next
    PushI Dy, Dr
Next
Dim FF$
    If Ff0 = "" Then
        FF = ValNm & " " & MapFunNN
    Else
        FF = Ff0
    End If
Stop
DrsMapItr = DrsFf(FF, Dy)
End Function

Function DrsRen(D As Drs, RenFf$) As Drs: DrsRen = Drs(FnyRen(D.Fny, RenFf), D.Dy): End Function
Function DrsSplitSS(D As Drs, SSCol$) As Drs
'Fm D     : It has a col @SSCol
'Fm SSCol : It is a col nm in @D whose value is SS.
'Ret  : a drs of sam ret but more rec by split@SSCol col to multi record
Dim I%: I = IxEle(D.Fny, SSCol)
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    Dim S: For Each S In Itr(SySs(Dr(I)))
        Dr(I) = S
        PushI Dy, Dr
    Next
Next
DrsSplitSS = Drs(D.Fny, Dy)
End Function

Function DwDist(A As Drs, CC$) As Drs:            DwDist = DrsFf(CC, DyWhDis(DrsSelFf(A, CC).Dy)): End Function
Function DwInsFf(A As Drs, FF$, NewDy()) As Drs: DwInsFf = Drs(SyAdd(FnyFF(FF), A.Fny), NewDy):    End Function
Function DyoSSy(Ssy$()) As Variant()
Dim Ss: For Each Ss In Itr(Ssy)
    Stop 'PushI DyoSSy, Tml(Ss)
Next
End Function

Function FmtLNewOGpno(Newl As Drs) As Drs
'@ NewL: L Ln NewL ! NewL may empty, when non-Emp, NewL <> Ln
'Ret D: L Ln NewL Gpno ! Gpno is running from 1:
'                      !   all conseq Ln with Emp-NewL is one group
'                      !   each non-Emp-NewL is one gp
Dim IGpno&, Dr, Dy(), Ln, NewL_, LasEmp As Boolean, Emp As Boolean

'For Each Dr In Itr(NewL.Dy)
'    PushI Dr, IsEmpty(Dr(2))
'    PushI Dy, Dr
'Next
'BrwDy Dy
'Erase Dy
'Stop
LasEmp = True
IGpno = 0
For Each Dr In Itr(Newl.Dy)
    Ln = Dr(1)
    NewL_ = Dr(2)
    Emp = IsEmpty(NewL_)
    If Not Emp Then If Ln = NewL_ Then Stop
    If IsEmpty(Ln) Then Stop
    Select Case True
    Case Not Emp: IGpno = IGpno + 1
    Case Emp And Not LasEmp: IGpno = IGpno + 1
    Case Else
    End Select
    PushI Dr, IGpno
    PushI Dy, Dr
    LasEmp = Emp
Next
FmtLNewOGpno = DrsFf("L Ln NewL Gpno", Dy)
End Function

Function FmtLNewOLines(NLn As Drs) As Drs
'Fm NLn: L Gpno NLn SNewL
'Ret Lines: L Gpno Lines
Dim Dr, L&, Gpno&, Lines$, NLn_$, SNewL
Dim Dy()
'Insp SNewL should have some Emp
'    Erase Dy
'    For Each Dr In NLn.Dy
'        PushI Dr, IsEmpty(Dr(2))
'        PushI Dy, Dr
'    Next
'    BrwDrs DrsFf("L Gpno NLn SNewL Emp", Dy)
'    Erase Dy
For Each Dr In Itr(NLn.Dy)
    AsgAy Dr, L, Gpno, NLn_, SNewL
    If IsEmpty(SNewL) Then
        Lines = NLn_
    Else
        Lines = NLn_ & vbCrLf & SNewL
    End If
    PushI Dy, Array(L, Gpno, Lines)
Next
FmtLNewOLines = DrsFf("L Gpno Lines", Dy)
'BrwDrs FmtLNewOLines: Stop
End Function

Function FmtLNewONLn(Gpno As Drs) As Drs
'@Gpno: L Ln NewL Gpno
'Ret E: L Gpno NLn SNewL ! NLn=L# is in front; SNewL = Spc is in front, only when nonEmp
Dim MaxL&: MaxL = EleMax(DcLngDrs(Gpno, "L"))
Dim NDig%: NDig = Len(CStr(MaxL))
Dim S$: S = Space(NDig + 1)
Dim Dy(), Dr, L&, Ln$, Newl, IGpno&, NLn$, SNewL
For Each Dr In Itr(Gpno.Dy)
    AsgAy Dr, L, Ln, Newl, IGpno
    NLn = AliR(L, NDig) & " " & Ln
    If IsEmpty(Newl) Then
        SNewL = Empty
    Else
        SNewL = S & Newl
    End If
    PushI Dy, Array(L, IGpno, NLn, SNewL)
Next
FmtLNewONLn = DrsFf("L Gpno NLn SNewL", Dy)
End Function

Function FmtNewOneG(NLn As Drs) As Drs
'@D: L Gpno NLn SNewL !
'Ret E: Gpno Lines ! Gpno now become uniq
Dim O$(), L&, LasG&, Dr, Dy(), Gpno&, NLn_$, SNewL
If NoRecDrs(NLn) Then Exit Function
LasG = NLn.Dy(0)(1)
For Each Dr In Itr(NLn.Dy)
    AsgAy Dr, L, Gpno, NLn_, SNewL
    If LasG <> Gpno Then
        PushI Dy, Array(Gpno, JnCrLf(O))
        Erase O
        LasG = Gpno
    End If
    PushI O, NLn_
    If Not IsEmpty(SNewL) Then PushI O, SNewL
Next
If Si(O) > 0 Then PushI Dy, Array(Gpno, JnCrLf(O))
FmtNewOneG = DrsFf("Gpno Lines", Dy)
End Function

Function FnyRen(Fny$(), RenFf$) As String()
Dim D As Dictionary: Set D = DiRenFf(RenFf)
Dim F: For Each F In Fny
    If D.Exists(F) Then
        PushI FnyRen, D(F)
    Else
        PushI FnyRen, F
    End If
Next
End Function

Function HasRecDrs(A As Drs) As Boolean: HasRecDrs = Si(A.Dy) > 0: End Function
Function HasRecDy(Dy()) As Boolean:       HasRecDy = Si(Dy) > 0:   End Function

Function IntoDrsC(Into, A As Drs, C)
Dim O, Ix%, Dy(), Dr
Ix = IxEle(A.Fny, C): If Ix = -1 Then Stop
O = Into
Erase O
Dy = A.Dy
If Si(Dy) = 0 Then IntoDrsC = O: Exit Function
For Each Dr In Dy
    Push O, Dr(Ix)
Next
IntoDrsC = O
End Function

Function IsEmpDrs(D As Drs) As Boolean
If HasRecDrs(D) Then Exit Function
If Si(D.Fny) > 0 Then Exit Function
IsEmpDrs = True
End Function

Function IsEqDrs(A As Drs, B As Drs) As Boolean
Select Case True
Case Not IsEqAy(A.Fny, B.Fny), Not IsEqAy(A.Dy, B.Dy)
Case Else: IsEqDrs = True
End Select
End Function

Function IsNeFf(A As Drs, FF$) As Boolean
IsNeFf = JnSpc(A.Fny) <> FF
End Function

Function IsSamDrEleCnt(A As Drs) As Boolean
IsSamDrEleCnt = IsSamDrEleCntDy(A.Dy)
End Function

Function IsSamDrEleCntDy(Dy()) As Boolean
If Si(Dy) = 0 Then IsSamDrEleCntDy = True: Exit Function
Dim C%: C = Si(Dy(0))
Dim Dr
For Each Dr In Itr(Dy)
    If Si(Dr) <> C Then Exit Function
Next
IsSamDrEleCntDy = True
End Function

Function RixDr&(Dy(), Dr)
Dim IDr, O&: For Each IDr In Itr(Dy)
    If IsEqAy(IDr, Dr) Then RixDr = O: Exit Function
    O = O + 1
Next
RixDr = -1
End Function

Function DrLas(D As Drs): DrLas = EleLas(D.Dy): End Function

Function RecLas(D As Drs) As Rec
Const CSub$ = CMod & "RecLas"
If Si(D.Dy) = 0 Then Thw CSub, "No RecLas", "Drs.Fny", D.Fny
RecLas = Rec(D.Fny, Av((EleLas(D.Dy))))
End Function

Function NDcDrs%(D As Drs):               NDcDrs = Max(Si(D.Fny), NDcDy(D.Dy)): End Function
Function NoRecDrs(D As Drs) As Boolean: NoRecDrs = NoRecDy(D.Dy):               End Function
Function NoRecDy(Dy()) As Boolean:       NoRecDy = Si(Dy) = 0:                  End Function

Function RLvlGpIx1(Dy()) As Long()

End Function

Function SelCol(Dy(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI SelCol, AwIxy(Dr, Ixy)
Next
End Function

Function SelDr(Dr(), IxyWiNeg&()) As Variant()
Dim Ix, U%: U = UB(IxyWiNeg)
For Each Ix In IxyWiNeg
    If IsBet(Ix, 0, U) Then
        PushI SelDr, Dr(Ix)
    Else
        PushI SelDr, Empty
    End If
Next
End Function

Function DySel(Dy(), Ciy) As Variant() ' Select Dy column: return a Dy which is SubSet-of-col of @Dy indicated by @SelIxy
If IsEmpAy(Ciy) Then DySel = Dy: Exit Function
Dim Dr: For Each Dr In Itr(Dy)
    PushI DySel, AwIxy(Dr, Ciy)
Next
End Function

Function SqDrs(D As Drs) As Variant()
If NoRecDrs(D) Then
    SqDrs = WSqFny(D.Fny)
    Exit Function
End If
Dim NC&, NR&, Dy(), Fny$()
    Fny = D.Fny
    Dy = D.Dy
    NC = Max(NDcDy(Dy), Si(Fny))
    NR = Si(Dy)
Dim O()
    ReDim O(1 To 1 + NR, 1 To NC)
    SetSqr O, 1, Fny      '<== Set O, R=1
    Dim R&: For R = 1 To NR
        SetSqr O, R + 1, Dy(R - 1) '<== Set O, Fm
    Next
SqDrs = O
End Function

Private Function WSqFny(Fny$()) As Variant()
Dim O()
ReDim O(1 To 2, 1 To Si(Fny))
Dim J%: For J = 0 To Si(Fny)
    SetSqr O, 1, Fny
Next
WSqFny = O
End Function

Sub UpdDbqCol_IfBlnk_ByPrv(D As Database, Q)
With Rs(D, Q)
    Dim L
    If Not .EOF Then L = .Fields(0).Value
    .MoveNext
    While Not .EOF
        If Trim(Nz(.Fields(0).Value, "")) = "" Then
            .Edit
            .Fields(0).Value = L
            .Update
        Else
            L = .Fields(0).Value
        End If
        .MoveNext
    Wend
    .Close
End With
End Sub

Function FstRecValDrsC(D As Drs, C)
Const CSub$ = CMod & "FstRecValDrsC"
If Si(D.Dy) = 0 Then Thw CSub, "No Rec", "Drs.Fny", D.Fny
FstRecValDrsC = D.Dy(0)(IxEle(D.Fny, C))
End Function


Private Sub B_DiCntDrs()
Dim Drs As Drs, Dic As Dictionary
'Drs = CVbe_Mth12Drs(CVbe)
Set Dic = DiCntDrs(Drs, "Nm")
BrwDi Dic
End Sub

Private Sub B_GpCol()
Dim DcDrs():            DcDrs = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
Dim RLvlGpIx&(): RLvlGpIx = Lngy(1, 1, 1, 3, 3, 2, 2, 3, 0, 0)
Dim G():                G = ColGp(DcDrs, RLvlGpIx)
Stop
End Sub

Private Sub B_GpDicDKG()
Dim Act As Dictionary, Dy(), Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Dy = Array(Dr1, Dr2, Dr3)
Set Act = DiGRxyToCy(Dy, Inty(0), 2)
Ass Act.Count = 2
Ass IsEqAy(Act("A"), Array(1, 2))
Ass IsEqAy(Act("B"), Array(3))
Stop
End Sub

Private Sub B_SelDistCnt()
'BrwDrs SelDistCnt(PFunDrs, "Mdn")
End Sub


Function DtDrs(D As Drs, Optional Dtn$ = "Dt") As Dt: DtDrs = Dt(Dtn, D.Fny, D.Dy): End Function
