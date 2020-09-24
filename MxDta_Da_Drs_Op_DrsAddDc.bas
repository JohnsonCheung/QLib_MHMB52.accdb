Attribute VB_Name = "MxDta_Da_Drs_Op_DrsAddDc"
Option Compare Database
Option Explicit
Function DrsAddDc(D As Drs, C$, V) As Drs
Dim Dy()
Dim Dr: For Each Dr In Itr(D.Dy)
    PushI Dr, V
    PushI Dy, Dr
Next
DrsAddDc = DrsFfAdd(D, C, Dy)
End Function

Function DrsAddDc2(A As Drs, F2$, V1, V2) As Drs
Dim Fny$(), Dy()
Fny = AyAdd(A.Fny, Tmy(F2))
Dy = DyAddDc2(A.Dy, V1, V2)
DrsAddDc2 = Drs(Fny, Dy)
End Function

Function DrsAddDc3(A As Drs, F3$, V1, V2, V3) As Drs
Dim Fny$(), Dy()
Fny = AyAdd(A.Fny, Tmy(F3))
Dy = DyAddDc3(A.Dy, V1, V2, V3)
DrsAddDc3 = Drs(Fny, Dy)
End Function

Function DrsAddDcFst(D As Drs, Gpcc$) As Drs
'Fm D    : ..@Gpcc.. ! a drs with col-@Gpcc
'Fm Gpcc :           ! col-@Gpcc in @D have dup.
'Ret     : @D Fst    ! a drs of col-Fst add to @D at end.  col-Fst is bool value.  TRUE when if it fst rec of a gp
'                    ! and rst of rec of the gp to FALSE
Dim O As Drs: O = DrsAddDc(D, "Fst", False) ' Add col-Fst with val all FALSE
If NoRecDrs(D) Then DrsAddDcFst = O: Exit Function
Dim GDy(): GDy = DrsSelFf(D, Gpcc).Dy  ' Dy with Gp-col only.
Dim R(): R = DyGp(GDy)                 ' Gp the @GDy into `DyGp`
Dim Cix&: Cix = UB(O.Dy(0))             ' Las col Ix aft adding col-Fst
Dim Rxy: For Each Rxy In R               ' for each gp, get the Row-ixy (pointing to @D.Dy)
    Dim Rix&: Rix = Rxy(0)               ' Rix is Row-ix pointing one of @D.Dy which is the fst rec of a gp
    O.Dy(Rix)(Cix) = True
Next
DrsAddDcFst = O
End Function

Function DrsAddDcGpno(D As Drs, NumColn$, GpnoColn$, Optional RunFmNum% = 1) As Drs
'Fm D : ..@NumColn..  ! must has a @NumColn which is a Num.  And assume they are sorted else thw
'Ret  : ..@GpnoColn  ! a drs with @GpnoColn added at end, which is a Gpno running from @RunFmNum
'                      if the conseq dr having @NumColn is in seg, given them a Gpno.
'                      Thw &IncIfJmp if @NumColn is not in ascending order.
Dim Gpno&: Gpno = RunFmNum
Dim Dy()
    If NoRecDrs(D) Then GoTo X
    Dim Ix%: Ix = IxEle(D.Fny, NumColn)
    Dim NumCur&
    Dim Dr: Dr = D.Dy(0)
    Dim NumLas&: NumLas = Dr(Ix)
    For Each Dr In Itr(D.Dy)
        NumCur = Dr(Ix)
        Gpno = W2IncIfJmp(Gpno, NumLas, NumCur)
        PushI Dr, Gpno
        PushI Dy, Dr
        NumLas = NumCur
    Next
X:
DrsAddDcGpno = DrsFfAdd(D, GpnoColn, Dy)
End Function

Private Function W2IncIfJmp(N&, NumLas, NumCur)
Const CSub$ = CMod & "W2IncIfJmp"
'Ret : Increased @N if NumLas has jumped else no chg @N
'      @N        if NumLas = NumCur or NumLas - 1 = CurNm
'      @N+1      If NumLas - 1 > NumCur
'      Otherwise Thw
Dim Dif&: Dif = NumCur - NumLas
Select Case Dif
Case 0, 1: W2IncIfJmp = N
Case Is > 1: W2IncIfJmp = N + 1
Case Else
    Thw CSub, "No in seq.  NumCur should > NumLas", "NumLas NumCur", NumLas, NumCur
End Select
End Function

