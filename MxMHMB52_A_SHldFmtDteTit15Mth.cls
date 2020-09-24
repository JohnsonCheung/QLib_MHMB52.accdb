VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldFmtDteTit15Mth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_SHld_FmtDteTit15Mth."
Private Function WWSampYm() As Ym: WWSampYm = Ym(19, 12):               End Function
Sub B_FmtSd():                                FmtSd WWSampWb, WWSampYm: End Sub
Sub B_FmtFc():                                FmtFc WWSampWb, WWSampYm: End Sub
Function Fmt(B As Workbook, A As Ym)
FmtFc B, A
FmtSd B, A
End Function

Private Function FmtFc(B As Workbook, A As Ym)
Dim Wsy() As Worksheet: Wsy = WWsyoFc(B)
Dim Sq(): Sq = WDT__Sq(A, 1)
Minvn B.Application   'Note: Without set visible & Max, R.Merge and R.UnMerge will break
Dim S: For Each S In Wsy
    Dim R As Range
    Set R = WDTRgFc(CvWs(S))
    R.UnMerge
    R.Value = Sq
    MgeTit R
Next
End Function
Sub B_WDTRgFc()
Dim B As Workbook: Set B = WWSampWb
Dim S As Worksheet: Set S = B.Sheets("Fc Stm")
Dim R As Range: Set R = WDTRgFc(S)
S.Select
S.Activate
R.Select
R.Activate
MsgBox R.Address
End Sub
Private Function WDTRgFc(FcWs As Worksheet) As Range: Set WDTRgFc = WDTRg(LoFst(FcWs), "M01", 15): End Function

Private Function FmtSd(B As Workbook, A As Ym)
Dim Sq(): Sq = WDT__Sq(A, -1, True)
Dim Wsy() As Worksheet: Wsy = WWsyoSd(B)
Minvn B.Application 'Without doing this R.Merge and R.UnMerge will break & doing Merge and Unmerge needs the Xls be visible
Dim S: For Each S In Wsy
    Dim R As Range
    Set R = WDTRgSd(CvWs(S))
    R.UnMerge
    R.Value = Sq
    MgeTit R
Next
End Function
Private Sub B_WDTRgSd()
Dim B As Workbook: Set B = WWSampWb
Dim S As Worksheet: Set S = B.Sheets("StkDays Stm")
S.Activate
MsgBox WDTRgSd(S).Address
End Sub
Private Function WDTRgSd(WsSd As Worksheet) As Range: Set WDTRgSd = WDTRg(LoFst(WsSd), "StkDays01", 15, 2, 1): End Function
Private Sub B_WDT__Sq()
Dim S As Ym: S = Ym(17, 4)
Dim N(): N = WDT__Sq(S)           'Normal
Dim P(): P = WDT__Sq(S, -1)       'Previous
Dim N2(): N2 = WDT__Sq(S, , True) 'Normal Double
Dim P2(): P2 = WDT__Sq(S, -1, True)   'Prevous Double
Stop
End Sub
Private Function WDT__Sq(M As Ym, Optional Ord As eSrt, Optional IsDouble As Boolean, Optional NMth% = 15) As Variant()
'@IsDouble: is the date title double column?
Dim Ym() As Ym: Ym = YMy(M, Ord, NMth)
Dim NC%: NC = NMth: If IsDouble Then NC = NC * 2
Dim O(): ReDim O(1 To 2, 1 To NC)
Dim C%, Ym1 As Ym, Yr%, Mmm$
If IsDouble Then
    For C = 1 To NMth
        Ym1 = Ym(C - 1)
        Yr = M.Y + 2000
        Mmm = MmmYm(M)
        O(1, C * 2 - 1) = Yr
        O(1, C * 2) = Yr
        O(2, C * 2 - 1) = Mmm
        O(2, C * 2) = Mmm
    Next
Else
    For C = 1 To NMth
        Ym1 = Ym(C - 1)
        O(1, C) = 2000 + M.Y
        O(2, C) = MmmYm(M)
    Next
End If
WDT__Sq = WDT_SetEmp_IfSamAs_PrvCol(O)
End Function
Private Function WDT_SetEmp_IfSamAs_PrvCol(Sq()) As Variant()
Dim O(): O = Sq
Dim R%: For R = 1 To UBound(Sq, 1)
    Dim C%: For C = UBound(Sq, 2) To 2 Step -1
        If O(R, C) = O(R, C - 1) Then O(R, C) = Empty
    Next
Next
WDT_SetEmp_IfSamAs_PrvCol = O
End Function

Private Function WDTRg(Lo As ListObject, Dte01ColNm$, NMth%, Optional NColPerMth% = 1, Optional NSpcRow%) As Range
':WDTRg: ! #Dte-title-range# is a 2-Rows-N-Months-range above DteTitAt having @NSpcRow
Dim R1%, R2%, C1%, C2%
R1 = -1 - NSpcRow
R2 = 0 - NSpcRow
C1 = 1
C2 = NMth * NColPerMth
Dim Dte01Cell As Range: Set Dte01Cell = CellLcHdr(Lo, Dte01ColNm)
Set WDTRg = RgRCRC(Dte01Cell, R1, C1, R2, C2)
End Function
Private Function WWSampWb() As Workbook
Dim T$: T = FxTmp("FmtFc"): ExpAttTpC "MH.SHldTp.Tp.xlsx", T
Debug.Print T
Set WWSampWb = WbFx(T)
End Function

Private Function WWsyoFc(B As Workbook) As Worksheet()
Dim S As Worksheet: For Each S In B.Sheets
    If HasPfx(S.Name, "Fc ") Then PushObj WWsyoFc, S
Next
End Function
Private Function WWsyoSd(B As Workbook) As Worksheet()
Dim S As Worksheet: For Each S In B.Sheets
    If HasPfx(S.Name, "StkDays ") Then PushObj WWsyoSd, S
Next
End Function
