Attribute VB_Name = "MxDta_Da_Dy_DyGp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dy_DyGp."

Function GpAsAyDy(D As Drs, Gpcc$) As Variant()
'@D : ..{Gpcc}..        ! it has col-@Gpcc
'Ret   : gp-of-:Dy:-of-@D ! each gp of dry has same @Gpcc val
Dim O()
    Dim SDy():  SDy = D.Dy              ' the src-dy to be gp
    Dim K As Drs: K = DrsSelFf(D, Gpcc)   ' a drs fm @D with grouping columns only
    Dim G():      G = DyGp(K.Dy)        ' gp-of-rxy-pointing-to-@D-row
    Dim Rxy: For Each Rxy In Itr(G)     ' Rxy is gp-of-rix-poiting-to-@D-row
        Dim ODy(): Erase ODy            ' Gp-of-@D-row with sam val of @Gpcc
        Dim Rix: For Each Rix In Rxy    ' Rix is rix-poiting-to-@D-row
            PushI ODy, SDy(Rix)         ' Push the @D-row to @ODy
        Next
        PushI O, ODy                    ' <-- put to @O, the output
    Next
GpAsAyDy = O
End Function

Function Gp(D As Drs, Gpcc$, Optional C$) As Drs
'@D : ..@Gpcc..@C.. ! it has col-Gpcc and optional col-C
'Ret   : @Gpcc #Gp ! where #Gp is opt gp of col-C, in :Av: @@
Dim OKey(), OGp()  ' Sam Si
    Dim SDy(): SDy = D.Dy               ' #Src-Dy. Source Dy to be gp
    Dim K As Drs: K = DrsSelFf(D, Gpcc)   ' #Key.    Only those @Gpcc column
    Dim IxGp%: IxGp = Si(K.Fny)         '          The column to put gp
    Dim KDr: For Each KDr In Itr(K.Dy)
        Dim Ix&: Ix = RixDr(OKey, KDr)
        If Ix = -1 Then
            PushI OGp, Array(SDy(Ix))
            PushI OKey, KDr
        Else
            PushI OGp(Ix), SDy(Ix)
        End If
    Next
Dim ODy()
Dim GpDr, J&: For Each GpDr In Itr(OKey)  ' For Each ele-of-OKey put corresponding ele-of-OGp at end to form a gp-rec
    PushI GpDr, OGp(J)               ' Put OGp(J) at end of #GpDr, now GpDr is a gp-rec
    PushI ODy, GpDr
    J = J + 1
Next
Gp = DrsAddDc(D, "Gp", ODy)
End Function

Function GRxyCiy(Dy(), Ciy) As Variant()
'Fm Dy : #Dta-Row-arraY# ! Dy to be gp.  It has all col as stated in @Ciy.
'Fm Ciy : #DcDrs-Ix-Array# ! Gpg which col of @Dy
'Ret    : Ay-of-Dy.  Each ele is a subset of @Dy in same gp.  @@
GRxyCiy = DyGp(DySel(Dy, Ciy)) ' sel the gp-ing col and gp it
End Function

Function DyGp(Dy()) As Variant()
'@Dy : :Dy        ! all col in @Dy will be used to gp
'Ret : :Av-of-Rxy ! #Av-of-Rxy# each ele in the returned @@Av is a :Rxy.  This :Rxy is pointing to @Dy so that all they are all eq @@
Dim K(), Dr, O(), Rix&: For Each Dr In Itr(Dy)
    Dim Gix&: Gix = RixDr(K, Dr)
    If Gix = -1 Then
        Dim Rxy&(): ReDim Rxy(0)
        Rxy(0) = Rix
        PushI O, Rxy      '<== Put Rix to Oup-O
        PushI K, Dr       '<-- Put Dr to K
    Else
        PushI O(Gix), Rix '<== Put Rix to Oup-O
    End If
    Rix = Rix + 1
Next
DyGp = O
End Function
Function DyGpCnt(Dy()) As Variant()
#If True Then
    DyGpCnt = WDyQuick(Dy)
#Else
    DyGpCnt = WDySlow(Dy)
#End If
End Function


Private Function WDyQuick(Dy()) As Variant()
End Function

Private Function WDySlow(Dy()) As Variant()
Const CSub$ = CMod & "WDySlow"
If Si(Dy) = 0 Then Exit Function
Dim OKeyDy(), OCnt&(), Dr
    Dim LasIx&: LasIx = Si(Dy(0))
    Dim J&
    For Each Dr In Dy
        If J Mod 500 = 0 Then Debug.Print "WDySlow"
        If J Mod 50 = 0 Then Debug.Print J;
        J = J + 1
        With IxOptDyDr(OKeyDy, Dr)
            Select Case .Som
            Case True: OCnt(.L) = OCnt(.L) + 1
            Case Else: PushI OKeyDy, Dr: PushI OCnt, 1
            End Select
        End With
    Next
    If Si(OKeyDy) <> Si(OCnt) Then Thw CSub, "Si Diff", "OKeyDy-Si OCnt-Si", Si(OKeyDy), Si(OCnt)
For J = 0 To UB(OCnt)
    PushI WDySlow, AyAdd(Array(OCnt(J)), OKeyDy(J)) '<===========
Next
End Function


Sub BrwGRxyCiyCntAy(Ay)
Brw JnGRxyCiyCntAy(Ay)
End Sub

Function GRxyCiyCntAy(Ay) As Variant()
If Si(Ay) = 0 Then Exit Function
Dim Dup, O(), X, T&, Cnt&
Dup = AwDup(Ay)
For Each X In Itr(Dup)
    Cnt = EleCnt(Ay, X)
    Push O, Array(X, EleCnt(Ay, X))
    T = T + Cnt
Next
Push O, Array("~Tot", T)
GRxyCiyCntAy = O
End Function

Function GRxyCiyCntAyWhDup(A) As Variant()
GRxyCiyCntAyWhDup = DyWhGt(GRxyCiyCntAy(A), 1, 1)
End Function

Function JnGRxyCiyCntAy(Ay) As String()
JnGRxyCiyCntAy = LyDy(GRxyCiyCntAy(Ay))
End Function

Private Sub B_JnGRxyCiyCntAy()
Dim Ay()
Brw JnGRxyCiyCntAy(Ay)
End Sub
