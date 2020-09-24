Attribute VB_Name = "MxDao_DrsPiv"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_DrsPiv."
Type KKCntMulItmColDy
    NKK As Integer
    Dy() As Variant
End Type
Enum eAgr: eSum: eCnt: eAvg: End Enum '#Aggregration#
Function DyBlk(Dy(), Kix%, Gix%) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim J%, O, K, Blk(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In Dy
    K = Dr(Kix)
    Gp = Dr(Gix)
    O_Ix = IxEle(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DyBlk = O
End Function
Function KKDrIx&(KKDr, FstColIsKKDrDy)
Dim Ix&, CurKKDr
For Each CurKKDr In Itr(FstColIsKKDrDy)
    If IsEqAy(KKDr, CurKKDr) Then
        KKDrIx = Ix
        Exit Function
    End If
    Ix = Ix + 1
Next
Ix = -1
End Function
Function KKDrToItmAyDualColDy(Dy(), KKColIx&(), ItmColIx&) As Variant()
Dim Dr, Ix&, KKDr(), Itm
Dim O() 'KKDr_To_ItmAy_DualColDy
For Each Dr In Itr(Dy)
    KKDr = AwIxy(Dy, KKColIx)
    Itm = Dr(ItmColIx)
    Ix = KKDrIx(KKDr, O)
    If Ix = -1 Then
        PushI O, Array(KKDr, Array(Itm))
    Else
        O(Ix)(1) = AyAdd(O(Ix)(1), Itm)
    End If
Next
KKDrToItmAyDualColDy = O
End Function
Function KKCntMulItmColDy(Dy(), KKColIx&(), ItmColIx&) As Variant()
Dim A(): A = KKDrToItmAyDualColDy(Dy, KKColIx, ItmColIx)
KKCntMulItmColDy = KKCntMulItmColDyD(A)
End Function
Function KKCntMulItmColDyD(KKDrToItmAyDualColDy()) As Variant()

End Function
Function GpDic(A As Drs, KK$, G$) As Dictionary
Dim Fny$()
Dim KeyIxy&(), Gix%
    Stop 'Fny = Tml(KK)
    KeyIxy = IxyEley(A.Fny, Fny)
    PushI Fny, G & "_Gp"
    Gix = IxEle(Fny, G)
Set GpDic = DiGRxyToCy(A.Dy, KeyIxy, Gix)
End Function

Function DiGRxyToCy(A, KeyIxy, G) As Dictionary
Const CSub$ = CMod & "DiGRxyToCy"
'If K < 0 Or G < 0 Then
'    Thw CSub, "K-Idx and G-Idx should both >= 0", "K-Idx G-Idx", K, G
'End If
'Dim Dr, U&, O As New Dictionary, KK, GG, Ay()
'U = UB(A): If U = -1 Then Exit Function
'For Each Dr In A
'    KK = Dr(K)
'    GG = Dr(G)
'    If O.Exists(KK) Then
'        Ay = O(KK)
'        PushI Ay, GG
'        O(KK) = Ay
'    Else
'        O.Add KK, Array(GG)
'    End If
'Next
'Set GRxyCiyDic = O
End Function

Function DrsFbt(Fb, T) As Drs: DrsFbt = DrsT(Db(Fb), T): End Function

Function DrsTSalTxt() As Drs
Stop 'DrsTSalTxt = DrsFx(MHOMB52.MHOMB52.FxiSalTxt, "Sheet1")
End Function
