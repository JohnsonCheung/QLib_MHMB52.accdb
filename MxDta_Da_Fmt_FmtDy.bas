Attribute VB_Name = "MxDta_Da_Fmt_FmtDy"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Fmt_FmtDy."
Private Type Pm
    W() As Integer
    BoolyAliR() As Boolean
    Zer As eZer
    QmkDta As Qmk
End Type
Private Sub B_FmtDy()
GoSub ZZ
Exit Sub
Dim Dy(), Fmt As eTblFmt, Zer As eZer, Wdt%, NoIx As Boolean, AlirCii$
ZZ:
    Dim O$()
    Erase O
    '
    Dim Sepln$: Sepln = String(100, "-")
    '
'    Dy = SampDy1: Fmt = eTblFmtTB: Zer = eZerShw: Wdt = 120: AlirCii = "": GoSub Push1
    Dy = sampDy2: Fmt = eTblFmtTb: Zer = eZerShw: Wdt = 120: AlirCii = "": GoSub Push1
'    Dy = SampDy3: Fmt = eTblFmtTB: Zer = eZerShw: Wdt = 120: AlirCii = "": GoSub Push1
'    PushI O, Sepln
'    Dy = SampDy1: Fmt = eTblFmtSS: Zer = eZerShw: Wdt = 120: AlirCii = "": GoSub Push1
'    Dy = SampDy2: Fmt = eTblFmtSS: Zer = eZerShw: Wdt = 120: AlirCii = "": GoSub Push1
'    Dy = SampDy3: Fmt = eTblFmtSS: Zer = eZerShw: Wdt = 120: AlirCii = "": GoSub Push1
    '
    Brw O
    Return
Push1:
    PushIAy O, FmtDy(Dy, Fmt, Zer, Wdt, AlirCii, NoIx)
    PushI O, ""
    Return
End Sub
Function FmtDy(Dy(), _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional Wdt%, _
Optional AlirCii$, _
Optional NoIx As Boolean) As String()
'(1) [Used by !FmtDrPms].  This fun is created because, !FmtDrPms will adjust the Wdty-of-Dy by Wdty-of-Fny.
'(2) [Sepln]             added after multiple lines or last Dr
If IsEmpAy(Dy) Then
    FmtDy = Sy("FmtDy: No record!")
    Exit Function
End If
Dim DyIxc()
Dim CiyAliR%()
Dim W%()
    DyIxc = DyAddDcIx(Dy, NoIx)
    CiyAliR = IntySs(AlirCii)
    If Not NoIx Then CiyAliR = AyItmAy(0, CiyAliR)
    W = WdtyDy(Dy, Wdt)
    If Not NoIx Then W = AyItmAy(NDig(UB(Dy) + 1), W)
Dim Lyy()
    Dim BoolyAliR() As Boolean: BoolyAliR = BoolyAliRCiy(CiyAliR, UB(W))
    Dim Q As Qmk: Q = QmkDta(Fmt)
    Dim Dr: For Each Dr In DyIxc
        PushI Lyy, FmtDrPm(Dr, W, Zer, Q, BoolyAliR)
    Next
Dim Sepln$: Sepln = SeplnWdty(W, Fmt)
FmtDy = W1_BdyLyy(Lyy, Sepln, Fmt)
End Function

Private Function W1BoolySepzLyy(Lyy()) As Boolean()
Dim O() As Boolean: ReDim O(UB(Lyy))
Dim J&, Ly: For Each Ly In Lyy
    If Si(Ly) > 1 Then
        O(J) = True
        If J > 0 Then O(J - 1) = True
    End If
    J = J + 1
Next
W1BoolySepzLyy = O
End Function
Private Function W1_BdyLyy(Lyy(), Sepln$, F As eTblFmt) As String()
Dim BoolySep() As Boolean: BoolySep = W1BoolySepzLyy(Lyy)
If F = eTblFmtTb Then PushI W1_BdyLyy, Sepln
Dim U&: U = UB(Lyy)
Dim Ly, J&: For Each Ly In Lyy
    PushIAy W1_BdyLyy, Ly
    If BoolySep(J) Then
        If J <> U Then
            PushI W1_BdyLyy, Sepln
        End If
    End If
    J = J + 1
Next
If F = eTblFmtTb Then PushI W1_BdyLyy, Sepln
End Function
