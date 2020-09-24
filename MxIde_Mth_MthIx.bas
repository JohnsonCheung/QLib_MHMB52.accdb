Attribute VB_Name = "MxIde_Mth_MthIx"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_MthIx."
Function BixMthFst&(Src$())
Dim I&: For I = 0 To UB(Src)
    If IsLnMth(Src(I)) Then BixMthFst = I: Exit Function
Next
BixMthFst = -1
End Function

Function Mthix&(Src$(), Mthn, Optional ShtMthTy$, Optional Fmix& = 0)
Dim I&: For I = Fmix To UB(Src)
    If HitMth(Src(I), Mthn, ShtMthTy) Then
        Mthix = I: Exit Function
    End If
Next
Mthix = -1
End Function

Private Sub B_HitMth()
Dim L$, Mthn, ShtMthTy$
GoSub T1
Exit Sub
T1:
    L = "Private Property Get W3AA$()"
    Mthn = "W3AA"
    ShtMthTy = "Get"
    Ept = True
    GoTo Tst
Tst:
    Act = HitMth(L, Mthn, ShtMthTy)
    C
    Return
End Sub
Function HitMth(L, Mthn, ShtMthTy$) As Boolean
Dim A As TMth: A = TMthL(L)
If Mthn <> A.Mthn Then Exit Function
If HitOptEq(A.ShtTy, ShtMthTy) Then HitMth = True: Exit Function
Debug.Print FmtQQ("HitMth: Mthn[?] Hits L but mis match given ShtMthTy[?].  Act ShtMthTy[?].  Ln=[?]", A.Mthn, A.ShtTy, ShtMthTy, L)
End Function

Function MthixFst&(Src$(), Optional Bix = 0)
For MthixFst = Bix To UB(Src)
    If IsLnMth(Src(MthixFst)) Then Exit Function
Next
MthixFst = -1
End Function

Function MthlnoFst&(M As CodeModule)
Dim J&: For J = 1 To M.CountOfLines
   If IsLnMth(M.Lines(J, 1)) Then
       MthlnoFst = J
       Exit Function
   End If
Next
End Function

Function MthySrcIx(Src$(), Mthix) As String(): MthySrcIx = AwBE(Src, Mthix, Mtheix(Src, Mthix)): End Function

Private Sub B_Mthlno()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MthnyM(A)
    For Each M In Ny
        DoEvents
        J = J + 1
        Push L, Mthlno(A, CStr(M))
        If J Mod 150 = 0 Then
            Debug.Print J, Si(Ny), "B_Mthlno"
        End If
    Next

    For Each Lno In L
        Push O, Lno & " " & A.Lines(Lno, 1)
    Next
BrwAy O
End Sub

Function MthixMN&(M As CodeModule, Mthn, Optional ShtMthTy$): MthixMN = MthixN(SrcM(M), Mthn, ShtMthTy): End Function
Function MthixN&(Src$(), Mthn, Optional ShtMthTy$)
Dim Ix&
For Ix = 0 To UB(Src)
    With TMthL(Src(Ix))
        If .Mthn = Mthn Then
            If HitOptEq(.ShtTy, ShtMthTy) Then
                MthixN = Ix
                Exit Function
            End If
            Debug.Print FmtQQ("MthixN: Given Mthn[?] Hit, not given ShtMthTy[?].  Act ShtMth[?]", Mthn, ShtMthTy, .ShtTy)
            If .ShtTy = "???" Then Stop
        End If
    End With
Next
MthixN = -1
End Function

Function HitOptEq(S, OptEq$) As Boolean ' If OptEq="" always return True, else return S=OptEq
If OptEq = "" Then HitOptEq = True: Exit Function
HitOptEq = S = OptEq
End Function

Function CMthlno&(): CMthlno = WMthLno(CMd, CLno): End Function
Private Function WMthLno&(M As CodeModule, LnoC&)
Dim L&: For L = LnoC To 1 Step -1
    If IsLnMth(M.Lines(L, 1)) Then WMthLno = L: Exit Function
Next
End Function

Function Mthlno&(M As CodeModule, MthNm, Optional FmLno& = 1)
Dim O&: For O = FmLno To M.CountOfLines
    If MthnL(M.Lines(O, 1)) = MthNm Then Mthlno = O: Exit Function
Next
End Function

Function CMthlcnt() As Lcnt: CMthlcnt = Mthlcnt(CMd, CMthn): End Function
Function Mthlcnt(M As CodeModule, Mthn, Optional FmLno& = 1) As Lcnt
Dim Lno&: Lno = Mthlno(M, Mthn, FmLno)
If Lno = 0 Then Exit Function
With Mthlcnt
    .Lno = Lno
    Dim E&: E = Mtheno(M, 1):    If E = 0 Then Thw CSub, "@Mthn has a Mthlno but no EnoSrcItm", "@Mthn Mthlno", Mthn, Lno
    .Cnt = E - Lno + 1
End With
End Function
