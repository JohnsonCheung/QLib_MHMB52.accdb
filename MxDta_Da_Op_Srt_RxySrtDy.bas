Attribute VB_Name = "MxDta_Da_Op_Srt_RxySrtDy"
Option Compare Binary ' See note in !WIsAyLe
Option Explicit
Const CMod$ = "MxDta_Da_Op_Srt_RxySrtDy."
Dim W1_Dy()
Dim W1_IsDesy() As Boolean
Private Type RxyLeGt: RxyLE() As Long: RxyGT() As Long: End Type ' See !W?RxyLeGt
Private Sub B_RxySrtDy()
Dim Dy(), IsDesy() As Boolean
GoSub T0
GoSub T1
Exit Sub
T0:
    Dy = DyVbl("2 a C|1 c B|3 b A")
    Ept = Lngy(1, 0, 2)
    Erase IsDesy
    GoTo Tst
T1:
    Dy = DyVbl("2 a C|1 c B|3 b A")
    Ept = Lngy(2, 0, 1)
    IsDesy = BoolyAp("t..")
    GoTo Tst
Tst:
    Act = RxySrtDy(Dy, IsDesy)
    C
    Return
End Sub

Function RxySrtDy(Dy(), IsDesy() As Boolean) As Long()
'UCol-@Dy and UB(@IsDesy) should be equal
If IsEmpAy(Dy) Then Exit Function
     W1_Dy = Dy
W1_IsDesy = IsDesy
RxySrtDy = WSrt(LngySno(Si(Dy)))
Erase W1_Dy
Erase W1_IsDesy
End Function

Private Function WSrt(Rxy&()) As Long()
Dim O&()
    Select Case UB(Rxy)
    Case -1
    Case 0: O = Rxy
    Case 1:
        O = WSwap(Rxy)
    Case Else
        Dim P&
            Dim C&(): C = Rxy  '*C :: Cur-Rxy
            P = Pop(C)         '*P :: Pivot-Rix, always use the Last ele of *C.  The Pivot-Part of @Rxy
        Dim L&()
        Dim H&()
            Dim M As RxyLeGt: M = WRxyLeGt(C, P)
            L = WSrt(M.RxyLE)      '*L :: Sorted-Low-Rxy of *C exl *P.  The Low-Part of @Rxy
            H = WSrt(M.RxyGT)      '*H :: High-Rxy of *C Exl *P.        The High-Part of @Rxy
        PushIAy O, L  ' Put Sorted-Low-Rxy
          PushI O, P  ' Put Pivot
        PushIAy O, H  ' Put Sorted-High-Rxy.  Put 3 parts of @Rxy together in L+P+H order means @Rxy is sorted
    End Select
WSrt = O
End Function

Private Function WRxyLeGt(Rxy&(), RixPiv&) As RxyLeGt ' Split @Rxy into @@RixLeGt by the @RixPiv
Dim RxyLE&(), RxyGT&()
    Dim DrPiv: DrPiv = W1_Dy(RixPiv)
    Dim Rix: For Each Rix In Rxy
        If WIsRixLe(Rix, DrPiv) Then
            PushI RxyLE, Rix
        Else
            PushI RxyGT, Rix
        End If
    Next
Dim O As RxyLeGt
O.RxyGT = RxyGT
O.RxyLE = RxyLE
WRxyLeGt = O
End Function
Private Function WIsAyLe(A, B) As Boolean ' Is Ay-@A less or equal to Ay-@B according the @W1_IsDesy
'@A, @B and @W_IsDesy should be same si
'For str ele,
'  if using > or <, it will be CasIgn, due to Option Compare Text
'  This module use Option Compare Binary!!
'  So the Ens3Opt needs to handle for this MxDtaDaSrtDyRxy
'  by using this approach, the sorting is always eCasSen
Dim J&: For J = 0 To UB(A)
    If W1_IsDesy(J) Then
        If A(J) < B(J) Then Exit Function
        If A(J) > B(J) Then WIsAyLe = True: Exit Function
    Else
        If A(J) > B(J) Then Exit Function
        If A(J) < B(J) Then WIsAyLe = True: Exit Function
    End If
Next
WIsAyLe = True
End Function
Private Function WIsRixLe(Rix, Dr) As Boolean: WIsRixLe = WIsAyLe(W1_Dy(Rix), Dr): End Function
Private Function WSwap(Ixy2&()) As Long()
Dim KeyB: KeyB = W1_Dy(Ixy2(1))
If WIsRixLe(Ixy2(0), KeyB) Then
    WSwap = Ixy2
Else
    PushI WSwap, Ixy2(1)
    PushI WSwap, Ixy2(0)
End If
End Function
