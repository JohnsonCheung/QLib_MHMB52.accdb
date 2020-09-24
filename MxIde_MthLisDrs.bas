Attribute VB_Name = "MxIde_MthLisDrs"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_MthLisDrs."
Public Const TFfMthLis$ = "Pjn CmpTy Mdn L Mdy Ty Mthn Tyc RetAs ShtPm"

Function DrsTMthLis() As Drs:                     DrsTMthLis = DrsTMthLisP(CPj):                       End Function
Function DrsTMthLisP(P As VBProject) As Drs:     DrsTMthLisP = DrsTMthLisMth(DrsTMthP(P)):             End Function
Function DrsTMthLisMth(DrsTMth As Drs) As Drs: DrsTMthLisMth = DrsSelFf(WAdd5Col(DrsTMth), TFfMthLis): End Function
Private Function WAdd5Col(WiMthln As Drs) As Drs
Dim ITyc   As Drs:     ITyc = W1_Tyc(WiMthln)
Dim IPm      As Drs:        IPm = W1_MthPm(ITyc)
Dim IShtPm   As Drs:     IShtPm = W1_ShtPm(IPm)
Dim IRetAs   As Drs:     IRetAs = W1_RetTyn(IShtPm)
WAdd5Col = W1_IsRetObj(IRetAs)
End Function
Private Function W1_MthPm(WiMthln As Drs, Optional IsDrp As Boolean) As Drs
W1_MthPm = DrsAddDcBetBkt(WiMthln, "Mthln:MthPm", IsDrp)
End Function
Private Function DrsAddDcBetBkt(D As Drs, ColnAs$, Optional IsDrp As Boolean) As Drs
Dim BetColn$, NewC$: AsgBrk1 ColnAs, ":", BetColn, NewC
If NewC = "" Then NewC = BetColn & "InsideBkt"
Dim Ix%: Ix = IxEle(D.Fny, BetColn)
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    PushI Dr, BetBkt(Dr(Ix))
    PushI Dy, Dr
Next
Dim O As Drs: O = DrsAddDc(D, NewC, Dy)
If IsDrp Then O = DrsDrpDc(O, BetColn)
DrsAddDcBetBkt = O
End Function

Private Function W1_IsRetObj(WiRetAs As Drs) As Drs
Dim IxRetAs%: IxRetAs = IxEle(WiRetAs.Fny, "RetAs")
Dim Dr, Dy(): For Each Dr In Itr(WiRetAs.Dy)
    Dim RetAs$: RetAs = Dr(IxRetAs)
    Dim R As Boolean: R = IsTynObj(RetAs)
    PushI Dr, R
    PushI Dy, Dr
Next
W1_IsRetObj = DrsAddDc(WiRetAs, "IsRetObj", Dy)
End Function
Private Function W1_RetTyn(WiMthln As Drs) As Drs
Dim I%: I = IxEle(WiMthln.Fny, "Mthln")
Dim Dr, Dy(): For Each Dr In Itr(WiMthln.Dy)
    Dim Mthln$: Mthln = Dr(I)
    Dim Ret$: Ret = W1_zRetTyn(Mthln)
    PushI Dr, Ret
    PushI Dy, Dr
Next
W1_RetTyn = DrsAddDc(WiMthln, "RetAs", Dy)
End Function
Private Function W1_zRetTyn$(Mthln)
Stop
End Function

Private Function W1_Tyc(WiMthln As Drs) As Drs
'Ret         : Add col-HasPm
Dim I%: I = IxEle(WiMthln.Fny, "Mthln")
Dim Dr, Dy(): For Each Dr In Itr(WiMthln.Dy)
    Dim Mthln$: Mthln = Dr(I)
    Dim Tyc$: Tyc = TycLn(Mthln)
    PushI Dr, Tyc
    PushI Dy, Dr
Next
W1_Tyc = DrsAddDc(WiMthln, "Tyc", Dy)
End Function

Private Function W1_ShtPm(WiMthPm As Drs) As Drs
'Ret         : Add col-ShtPm
Dim I%: I = IxEle(WiMthPm.Fny, "MthPm")
Dim Dr, Dy(): For Each Dr In Itr(WiMthPm.Dy)
    Dim Mthpm$: Mthpm = Dr(I)
    Dim ShtPm1$: ShtPm1 = ShtMthpm(Mthpm)
    PushI Dr, ShtPm1
    PushI Dy, Dr
Next
W1_ShtPm = DrsAddDc(WiMthPm, "ShtPm", Dy)
End Function

Private Function DwRetAy(WiRetAs As Drs, RetAy As eTri) As Drs
If RetAy = eTriOpn Then DwRetAy = WiRetAs: Exit Function
Dim RetAy1 As Boolean: RetAy1 = BoolTri(RetAy)
Dim IRetAs%: IRetAs = IxEle(WiRetAs.Fny, "RetAs")
Dim ODy()
    Dim Dr: For Each Dr In Itr(WiRetAs.Dy)
        Dim RetAs$: RetAs = Dr(IRetAs)
        If HasSfx(RetAs, "()") = RetAy1 Then PushI ODy, Dr
    Next
DwRetAy = Drs(WiRetAs.Fny, ODy)
End Function

Private Function HasAp(Mthpm) As Boolean
Dim A$(): A = SplitCmaSpc(Mthpm): If Si(A) = 0 Then Exit Function
HasAp = HasPfx(EleLas(A), "ParamArray ")
End Function

Private Function DwAnyAp(WiMthPm As Drs, HasAp0 As eTri) As Drs
If HasAp0 = eTriOpn Then DwAnyAp = WiMthPm: Exit Function
Dim HasAp1 As Boolean: HasAp1 = BoolTri(HasAp0)
Dim IMthPm%: IMthPm = IxEle(WiMthPm.Fny, "MthPm")
Dim ODy()
    Dim Dr: For Each Dr In Itr(WiMthPm.Dy)
        Dim Mthpm$: Mthpm = Dr(IMthPm)
        If HasAp1 = HasAp(Mthpm) Then PushI ODy, Dr
    Next
DwAnyAp = Drs(WiMthPm.Fny, ODy)
End Function

Private Function DwNPm(D As Drs, NPm%) As Drs
If NPm < 0 Then DwNPm = D: Exit Function
Dim Ix%: Ix = IxEle(D.Fny, "MthPm")
Dim ODy(), Dr, Pm$: For Each Dr In Itr(D.Dy)
    Pm = Dr(Ix)
    If Si(SplitCma(Pm)) = NPm Then PushI ODy, Dr
Next
DwNPm = Drs(D.Fny, ODy)
End Function
