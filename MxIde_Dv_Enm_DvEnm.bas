Attribute VB_Name = "MxIde_Dv_Enm_DvEnm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Deri_Enm_Dvenm."
Private Type TDvenm
    Cdl As String          ' Is LinesApLn of CtorMthl, AyMthl UdtlOpt
    MthnyDlt() As String      '
    Cnstly() As String
End Type
Private Sub B_Dvenm():    DvEnmM CMd:                  End Sub
Private Sub B_DvenmMdn(): DvEnmMdn "MxDao_Db_Schm_Ud": End Sub
Private Sub B_Srcdvenm()
GoSub T1
Exit Sub
Dim Src$()
T1:
    Src = SrcMdn("MxDao_Db_LnkImp")
    GoTo Tst
Tst:
    Act = SrcDvenm(Src)
    If IsEqAy(Src, Act) Then
        MsgBox "Same"
    Else
        MsgBox "Diff"
        Vc Src, "Bef"
        Vc Act, "Aft"
    End If
    Return
End Sub
Private Sub B_W1_TDvenmSrcU()
GoSub ZZ
Exit Sub
ZZ:
    Dim Src$()
    Dim O$(), U() As TEnm: 'U = Enm_UdyPC
    Dim J%: 'For J = 0 To UbTEnm(U)
        GoSub Push
    'Next
    BrwLsy O
    Exit Sub
Push:
    With W1_TDvenmSrcU(Src, U(J))
        PushNB O, LinesApLn(JnSpc(.MthnyDlt), .Cdl)
    End With
    Return
End Sub
Private Sub B_X8_Ctor()
Const T As Boolean = True
Const F As Boolean = False
GoSub T1
Exit Sub
Dim Src$(), IsPrv, EnmnLn$, Mbr() As TEmbr, GenEnmm4Fun As Boolean, GenEnmt As Boolean
T1:
    IsPrv = True
    EnmnLn = "eXXsdf"
    Erase Mbr
    GenEnmm4Fun = True
    GenEnmt = True
    Ept = RplVBar("Private Function sdf(AA() As ABC, BB As TEnm, CC As Ws, DD%()) As sdf|With sdf" & _
        "|    .AA = AA" & _
        "|    .BB = BB" & _
        "|    Set .CC = CC" & _
        "|    .DD = DD" & _
        "|End With|End Function")
    GoTo Tst
Tst:
    Dim U As TEnm
    'U = Enm(IsPrv, EnmnLn, Mbr, GenEnmm4Fun, GenEnmt)
    Act = W1CdldvenmSrcU(Src, U)
    C
    Return
End Sub
Private Sub B_W1CdlDvenmfunU()
GoSub T1
Dim E As TEnm
Exit Sub
T1:
    'E = Enm()
    Ept = ""
'Function EnmmXlsTy$(): EnmmXlsTy = "eXlsTy? Nbr Txt TorN Dte Bool": End Function
'Function EnmvXlsTy(S$) As eXlsTy
'Const CSub$ = CMod & "EnmXlsTyNm"
'Dim I%: I = IxEle(EnmsyXlsTy, S)
'If I = -1 Then ThwEnm CSub, S, EnmmXlsTy
'End Function
'Function EnmsyXlsTy() As String()
'Static Ny$(): If Si(Ny) = 0 Then Ny = SplitSpc(EnmmXlsTy)
'EnmsyXlsTy = Ny
'End Function
'Function EnmsXlsTy$(E As eXlsTy): EnmsXlsTy = EnmsyXlsTy()(E): End Function
    GoTo Tst
Tst:
    Act = W1CdlDvenmfunU(E)
    Stop
    C
    Return
End Sub
Sub DvEnmMC():                       DvEnmM CMd:                                End Sub
Sub DvEnmMdn(Mdn):                   DvEnmM Md(Mdn):                            End Sub
Sub WrtMsrcEnmP(P As VBProject):     WrtMsrcy MsrcyDvEnmP(P):                   End Sub
Private Sub DvEnmM(M As CodeModule): RplMdSrclopt M, SrcoptDvenmzCmp(M.Parent): End Sub
Function SrcoptDvenmzCmp(C As VBComponent) As Stropt
Dim S$(): S = SrcCmp(C): If Si(S) = 0 Then Exit Function
With WSrcloptDvenmSrc(S)
    If Not .Som Then Exit Function
    SrcoptDvenmzCmp = SomStr(LinesMix(JnCrLf(S), .Str))
End With
End Function
Function SrcDvenm(Src$()) As String()
Dim U() As TEnm: U = TEnmyDcl(DclSrc(Src))
Dim O$(): O = Src
Dim J%: For J = 0 To UbTEnm(U)
    O = WSrcDvenmU(O, U(J))
Next
SrcDvenm = O
End Function
Private Function WSrcloptDvenmSrc(Src$()) As Stropt: WSrcloptDvenmSrc = StroptOldNew(JnCrLf(Src), WSrcldvenmSrc(Src)): End Function
Private Function WSrcDvenmU(Src$(), U As TEnm) As String()
With W1_TDvenmSrcU(Src, U)
WSrcDvenmU = SrcEnsMth(Src, .Cdl, .MthnyDlt)
End With
End Function
Private Function WSrcldvenmSrc$(Src$()): WSrcldvenmSrc = JnCrLf(SrcDvenm(Src)): End Function
Private Function W1_TDvenmSrcU(Src$(), U As TEnm) As TDvenm
If Not WIsGenTEnm(U) Then Exit Function
Dim O As TDvenm
O.MthnyDlt = W1MthnydltTEnm(U)
O.Cdl = W1CdldvenmSrcU(Src, U)
W1_TDvenmSrcU = O
End Function
Private Function W1CdldvenmSrcU$(Src$(), U As TEnm) '#Cdl-For-Cur-Udt#
Dim O$()
With U
  If .IsGenEnmm4Fun Then PushNB O, W1CdlDvenmfunU(U)
  If .IsGenEnmt Then PushNB O, X8_Enmt(Src, U)
End With
W1CdldvenmSrcU = JnCrLf(O)
End Function
Private Function WIsGenTEnm(U As TEnm) As Boolean
Select Case True
Case U.IsGenEnmm4Fun, U.IsGenEnmt: WIsGenTEnm = True
End Select
End Function
Private Function W1MthnydltTEnm(U As TEnm) As String()
Dim N$: N = RmvFst(U.Enmn)
Dim O$()
    With U
    If .IsGenEnmm4Fun Then PushIAy O, NyQtp2("?" & U.Enmn, "Enms Enmsy Enmm Enmv")
      If .IsGenEnmt Then PushI O, "Enmt" & N
    End With
W1MthnydltTEnm = O
End Function
Private Function W1CdlDvenmqssU$(U As TEnm)

End Function

Private Function W1CdlDvenmfunU$(U As TEnm)
'Function EnmmXlsTy$(): EnmmXlsTy = "eXlsTy? Nbr Txt TorN Dte Bool": End Function
'Function EnmvXlsTy(S$) As eXlsTy
'Const CSub$ = CMod & "EnmXlsTyNm"
'Dim I%: I = IxEle(EnmsyXlsTy, S)
'If I = -1 Then ThwEnm CSub, S, EnmmXlsTy
'End Function
'Function EnmsyXlsTy() As String()
'Static Ny$(): If Si(Ny) = 0 Then Ny = SplitSpc(EnmmXlsTy)
'EnmsyXlsTy = Ny
'End Function
'Function EnmsXlsTy$(E As eXlsTy): EnmsXlsTy = EnmsyXlsTy()(E): End Function
      Const C_Enmv$ = "?Function Enmv?$() As e?|Const CSub$ = CMod & ""Enmv?""|Dim I%: I = IxEle(Enmsy?, S)|If I = -1 Then ThwEnm CSub, S, Enmm?|End Function"
      Const C_EnmqssQtp2$ = "?Function Enmv?$() As e?|Const CSub$ = CMod & ""Enmv?""|Dim I%: I = IxEle(Enmsy?, S)|If I = -1 Then ThwEnm CSub, S, Enmm?|End Function"
      Const C_EnmqssSs$ = "?Function Enmv?$() As e?|Const CSub$ = CMod & ""Enmv?""|Dim I%: I = IxEle(Enmsy?, S)|If I = -1 Then ThwEnm CSub, S, Enmm?|End Function"
     Const C_Enmsy$ = "?Function Enmsy?() As String()|Static Ny$(): If Si(Ny) = 0 Then Ny = SplitSpc(Enmm?)|Enmsy? = Ny|End Function"
      Const C_Enms$ = "?Function Enms?$(E As e?): Enms? = Enmsy?()(E): End Function"
Dim P$: P = MdyPrv(U.IsPrv) ' PrvPfx
Dim N$: N = RmvFst(U.Enmn)
Dim O$()
Dim StrEnmm$: StrEnmm = X9_StrEnmm(U)
Dim C_Enmqss$: C_Enmqss = "": Stop
PushI O, FmtQQ(C_Enmqss, P, N, N, StrEnmm)
PushI O, FmtQQ(C_Enmv, P, N, N, StrEnmm)
PushI O, FmtQQ(C_Enmsy, P, N, N, StrEnmm)
PushI O, FmtQQ(C_Enms, P, N, N, StrEnmm)
W1CdlDvenmfunU = JnCrLf(O)
End Function
Private Function X8_Enmt$(Src$(), U As TEnm)
  Const C_OptCtor$ = "Function Opt?(Som, A As ?) As ?Opt: With ?Opt: .Som = Som: .? = A: End With: End Function"
   Const C_OptSom$ = "Function Som?(A As ?) As ?Opt: Som?.Som = True: Som?.? = A: End Function"
  Const C_OptPush$ = "Sub PushOpt?(A() As ?, M As ?Opt)|With M|   If .Som Then Push? A, .?|End With|End Sub"
Dim N$: N = U.Enmn
Dim P$: P = MdyPrv(U.IsPrv)
Dim A$: A = P & RplQ(C_OptCtor, N)
Dim B$: B = P & RplQ(C_OptSom, N)
Dim C$: C = P & RplQ(RplVBar(C_OptPush), N)
X8_Enmt = LinesApLn(A, B, C)
End Function
Private Function X9_StrEnmm$(U As TEnm)
Dim Ny$(): Ny = MbnyTEnm(U)
Dim N$: N = RmvFst(U.Enmn)
X9_StrEnmm = N & "? " & JnSpc(AmRmvPfx(Ny, N))
End Function

