Attribute VB_Name = "MxDta_Da_GpCnt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_GpCnt."
Type GpCnt
    GpDy() As Variant ' Gp-Dy
    Cnt() As Long
End Type

Function GpCntDy(Dy()) As GpCnt
'@D : :Drs-..{Gpcc}    ! it has columns-Gpcc
'Ret  :GpCnt  ! each ele-of-:GpCnt.Gp is a dr with fld as described by @Gpcc.  :GpCnt.Cnt is rec cnt of such gp
Dim OGpDy(), OCnt&()
    Dim Dr: For Each Dr In Itr(Dy)
        Dim Ix&: Ix = RixDr(OGpDy, Dr)
        If Ix = -1 Then
            PushI OCnt, 1
            PushI OGpDy, Dr
        Else
            OCnt(Ix) = OCnt(Ix) + 1
        End If
    Next
GpCntDy.GpDy = OGpDy
GpCntDy.Cnt = OCnt
End Function

Function GpCntFny(D As Drs, GpFny$()) As GpCnt
'@D : :Drs-..{Gpcc}    ! it has columns-Gpcc
'Ret  :GpCnt  ! each ele-of-:GpCnt.Gp is a dr with fld as described by @Gpcc.  :GpCnt.Cnt is rec cnt of such gp
Dim OGpDy(), OCnt&()
    Dim A As Drs: A = DrsSelFny(D, GpFny)
    Dim I%: I = Si(A.Fny)
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim Ix&: Ix = RixDr(OGpDy, Dr)
        If Ix = -1 Then
            PushI OCnt, 1
            PushI OGpDy, Dr
        Else
            OCnt(Ix) = OCnt(Ix) + 1
        End If
    Next
GpCntFny.GpDy = OGpDy
GpCntFny.Cnt = OCnt
End Function
Function GpCntAllCol(D As Drs) As GpCnt
GpCntAllCol = GpCntDy(D.Dy)
End Function
Function GpCnt(D As Drs, Gpcc$) As GpCnt
GpCnt = GpCntFny(D, SySs(Gpcc))
End Function
