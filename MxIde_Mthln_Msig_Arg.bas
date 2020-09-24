Attribute VB_Name = "MxIde_Mthln_Msig_Arg"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_MSig_Arg."
#If Doc Then
'Cml
' Arg TArgment
' Argn TArgment name
'Tm
' Arg::S Comma term of method parameter
'
'Enum eArgM ! TArgment modifier, which all string before Argn
'
' :    :S ! Argm Nm ArgSfx Dft
':ShtArgm: :S ! One-of-:ShtArgmAy
'            ! :Sfx: ::  Tyc[Bkt] | vbColon AsTy
'            ! :Dft: :: [ChrEq StrDft] !
':TycFun: :C #Mth-Ty-Chr# ! one of :TycLis
':TycLis: :S #Mth-Ty-Chr-List# ! !@#$%^&
':C:  :Chr #Char# ! One single char
':Tyc: :TycFun:
#End If
Public Const LisTyc$ = "!@#$%^&"
Function Argn$(Arg):       Argn = TakNm(RmvArgm(Arg)): End Function
Function RmvArgm$(Arg): RmvArgm = RmvPfxy(Arg, Argmy): End Function

Function ShtArgyPC() As String()
Dim Arg: For Each Arg In AwDis(ArgyPC)
    PushI ShtArgyPC, ShtArg(Arg)
Next
ShtArgyPC = AySrtQ(ShtArgyPC)
End Function
Function ShtArgTArg$(A As TArg): ShtArgTArg = X_ShtArg(A): End Function

Private Sub B_ShtArg()
GoSub ZZ
GoSub T0
Exit Sub
Dim Arg
T0:
     Arg = "Optional UseVc As Boolean"
     Ept = "?UseVc?"
     GoTo Tst
Tst:
    Act = ShtArg(Arg)
    C
    Return
ZZ:
    Dim S() As S12
    Dim A: For Each A In ArgyPC
        PushS12 S, S12(A, ShtArg(A))
    Next
    BrwS12y S
    Return
End Sub
Private Function W3S12yPm(Argy$()) As S12()
Dim Arg: For Each Arg In Itr(Argy)
    PushS12 W3S12yPm, S12(Arg, ShtArg(Arg))
Next
End Function
Function ShtArg$(Arg): ShtArg = X_ShtArg(TArgArg(Arg)): End Function

Private Function X_ShtArg$(A As TArg) ' Will have no spc in ShtArg
Dim Pfx$: Pfx = WPfx(A.Argm)
Dim Sfx$: Sfx = VsfxTVt(A.Vt)
Dim Dft$: Dft = StrTrue(A.Dft <> "", "=" & A.Dft)
      X_ShtArg = Pfx & A.Argn & Sfx & Dft
    If X_ShtArg = "Optional" Then Stop
End Function
Private Function WPfx$(M As eArgm)
Const CSub$ = CMod & "WPfx"
Dim O$
Select Case True
Case M = eArgmRef: O = ""
Case M = eArgmVal: O = "*"
Case M = eArgmRefOpt: O = "?"
Case M = eArgmValOpt: O = "?*"
Case M = eArgmAp: O = ".."
Case Else: ThwEnm CSub, M, EnmmArgm
End Select
WPfx = O
End Function

Private Sub B_Argn()
Dim Arg$
'GoSub T1
GoSub YY
Exit Sub
T1:
    Arg = "Optional Fnn$"
    Ept = "Fnn"
    GoTo Tst
Tst:
    Act = Argn(Arg)
    C
    Return
YY:
Dim O() As S12
Dim A: For Each A In ArgyP(CPj)
    PushS12 O, S12(Argn(Arg), Arg)
Next
BrwS12y O
End Sub


Private Sub B_ShtArgy()
Dim A() As S12
Dim Arg: For Each Arg In AwDis(ArgyPC)
    PushS12 A, S12(Arg, ShtArg(Arg))
Next
BrwS12y A
End Sub

Function ShtArgy(Argy$()) As String()
Dim A: For Each A In Itr(Argy)
    PushI ShtArgy, ShtArg(A)
Next
End Function

Function RmvTyc$(S):             RmvTyc = RmvfstChrInLis(S, LisTyc): End Function
Function Argm$(Arg):               Argm = PfxPfxySpc(Arg, Argmy):    End Function '#Argm:TArgment-Modifier#
Function ArgnyPC() As String(): ArgnyPC = ArgnyP(CPj):               End Function
Function ArgnyP(P As VBProject) As String()
Dim O$()
    Dim Mthln: For Each Mthln In MthlnyP(P)
        PushIAy O, Argy(Mthln)
    Next
O = AwDis(O)
ArgnyP = AySrtQ(O)
End Function

Function ShtPm$(Mthpm)
Dim O$()
Dim Arg: For Each Arg In Itr(SplitCmaSpc(Mthpm))
    PushI O, ShtArg(Arg)
Next
ShtPm = JnSpc(O)
End Function

Function VsfxArg$(Arg$):   VsfxArg = RmvNm(ArgItm(Arg)):                                                  End Function
Function ShtVtArg$(Arg$): ShtVtArg = ShtVsfx(VsfxArg(ArgItm(Arg))):                                       End Function
Function ArgItm$(Arg):      ArgItm = BefOrAll(RmvPfxSpc(RmvPfxSpc(Arg, "Optional"), "ParamArray"), " ="): End Function

Function FmtPm(Pm$, Optional IsNoBkt As Boolean) 'Pm is wo bkt.
Dim A$: A = Replace(Pm, "Optional ", "?")
Dim B$: B = Replace(A, " As ", ":")
Dim C$: C = Replace(B, "ParamArray ", "...")
If IsNoBkt Then
    FmtPm = C
Else
    FmtPm = QuoSq(C)
End If
End Function

Function RetAsDclSfx$(DclSfx)
If DclSfx = "" Then Exit Function
Dim B$
Dim F$: F = ChrFst(DclSfx)
If IsTyc(F) Then
    If Len(DclSfx) = 1 Then Exit Function
    B = RmvFst(DclSfx): If B <> "()" Then Stop
    RetAsDclSfx = " As " & Tycn(F) & "()"
    Exit Function
End If
If TycTycn(DclSfx) <> "" Then Exit Function
Select Case True
Case Left(DclSfx, 4) = " As ":   RetAsDclSfx = DclSfx
Case Left(DclSfx, 6) = "() As ": RetAsDclSfx = Mid(DclSfx, 3) & "()"
Case DclSfx = "()":              RetAsDclSfx = " As Variant()"
Case Else: Stop
End Select
End Function
Function TycDclSfx$(DclSfx)
If Len(DclSfx) = 1 Then
    If IsTyc(DclSfx) Then TycDclSfx = DclSfx
End If
End Function

Function ArgSfxRet$(Ret)
'Ret is either FunRetTyc (in Sht-Tyc) or
'              FunRetAs    (The Ty-Str without As)
Select Case True
Case IsTyc(ChrFst(Ret)): ArgSfxRet = Ret
Case HasSfx(Ret, "()") And TycTycn(RmvSfx(Ret, "()")) <> "": ArgSfxRet = TycTycn(RmvSfx(Ret, "()")) & "()"
Case Else: ArgSfxRet = " As " & Ret
End Select
End Function
Function ArgSfx$(Arg)
Const CSub$ = CMod & "ArgSfx"
Dim L$: L = RmvArgm(Arg)
Dim Nm$: Nm = ShfNm(L): If Nm = "" Then Thw CSub, "Given Arg is invalid (No name aft arg modifier)", "Arg", Arg
ArgSfx = BefOrAll(L, " = ")
End Function
Function ArgSfxy(Argy$()) As String()
Dim Arg: For Each Arg In Itr(Argy)
    PushI ArgSfxy, ArgSfx(Arg)
Next
End Function

Function Argy(Mthln) As String():   Argy = SplitCmaSpc(BetBkt(Mthln)): End Function
Private Sub B_ArgyPC():                    VcAy ArgyPC:                End Sub
Function ArgyPC() As String():    ArgyPC = ArgyP(CPj):                 End Function
Function ArgyP(P As VBProject) As String()
Dim A$()
    Dim Mthln: For Each Mthln In Itr(MthlnySrc(SrcP(P)))
        PushIAy A, Argy(Mthln)
    Next
ArgyP = AwDis(A)
End Function

Function Argny(A() As TArg) As String()
Dim J%: For J = 0 To UbTArg(A)
    PushI Argny, A(J).Argn
Next
End Function

Function ArgnyPm(Pm$) As String()
Dim Ay$(): Ay = SplitCmaSpc(Pm)
Dim I: For Each I In Itr(Ay)
    PushI ArgnyPm, TakNm(I)
Next
End Function

Function TArgyP(Pm$) As TArg() ' @Pm is the Bet-Bkt-of a Mthln
Dim A: For Each A In Itr(SplitCmaSpc(Pm))
    PushTArg TArgyP, TArgArg(A)
Next
End Function

Function TArgArg(Arg) As TArg
Const CSub$ = CMod & "TArgArg"
Dim S$:                 S = Trim(Arg)
Dim Mdy As eArgm:     Mdy = ShfeArgm(S)
Dim Nm$:               Nm = ShfNm(S): If Nm = "" Then Thw CSub, "Invalid Arg: no name", "Arg", Arg
With Brk1(S, "=")
    Dim T As TVt: T = TVtVsfx(.S1)
    Dim Dft$: Dft = .S2
    TArgArg = TArg(Mdy, Nm, T, Dft)
End With
End Function

Function ShfeArgm(OArg$) As eArgm: ShfeArgm = eArgm(ShfPfxySpc(OArg, Argmy, "ByRef")): End Function

Private Sub B_ShtArg1()
Dim O$()
Dim A: For Each A In ArgyPC
    PushI O, ShtArg(A)
Next
VcAy AySrtQ(O)
End Sub
