Attribute VB_Name = "MxTp_IndItm"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_IndItm."
Type TInd: Hdr As String: Ix As Long: Ly() As String: End Type 'Deriving(Ctor Ay)
Function TInd(Hdr, Ix, Ly$()) As TInd
With TInd
    .Hdr = Hdr
    .Ix = Ix
    .Ly = Ly
End With
End Function
Function AddTInd(A As TInd, B As TInd) As TInd(): PushTInd AddTInd, A: PushTInd AddTInd, B: End Function
Sub PushTIndAy(O() As TInd, A() As TInd): Dim J&: For J = 0 To TIndUB(A): PushTInd O, A(J): Next: End Sub
Sub PushTInd(O() As TInd, M As TInd): Dim N&: N = TIndSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function TIndSI&(A() As TInd): On Error Resume Next: TIndSI = UBound(A) + 1: End Function
Function TIndUB&(A() As TInd): TIndUB = TIndSI(A) - 1: End Function

Private Sub B_TIndSrc()
Dim A() As TInd: A = TIndSrc(TIndSrcPC)
Stop
End Sub
Function TIndSrc(SrcInd$()) As TInd()
Const CSub$ = CMod & "TIndSrc"
If Si(SrcInd) = 0 Then Thw CSub, "SrcInd is empty"
Dim A$(): A = WRmvD3(SrcInd$())
If IsLnInd(A(0)) Then Thw CSub, "First chr of first line of SrcInd must not be blank", "SrcInd aft rmv D3", A
TIndSrc = WTIndy(A)
End Function
Private Function WRmvD3(TpLy$()) As String() ' Rmv all D3ln and D3Ssub
Dim L: For Each L In Itr(TpLy)
    If HasSsub(L, "---") Then
        With Brk1(L, "---", NoTrim:=True)
            If Trim(.S1) <> "" Then
                PushI WRmvD3, .S1
            End If
        End With
    Else
        PushI WRmvD3, L
    End If
Next
End Function
Private Function WTIndy(SrcInd$()) As TInd()
WTIndy = WTIndyzH(WHixy(SrcInd), SrcInd)
End Function
Private Function WHixy(SrcInd$()) As Long()
Dim Ix&: For Ix = 0 To UB(SrcInd)
    Dim L$: L = SrcInd(Ix)
    If Not IsLnInd(L) Then PushI WHixy, Ix
Next
End Function
Private Function WTIndyzH(Hixy&(), SrcInd$()) As TInd()
Dim H&(): H = Hixy: PushI H, Si(SrcInd)
Dim Ix&: For Ix = 0 To UB(Hixy)
    Dim B&: B = H(Ix) + 1
    Dim E&: E = H(Ix + 1) - 1
    Dim Ly$(): Ly = AmLTrim(AwBE(SrcInd, B, E))
    PushTInd WTIndyzH, TInd(SrcInd(B), B, Ly)
Next
End Function
