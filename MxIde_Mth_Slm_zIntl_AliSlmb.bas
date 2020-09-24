Attribute VB_Name = "MxIde_Mth_Slm_zIntl_AliSlmb"
Option Compare Text
Const CMod$ = "MxIde_Mth_Slmb_AliSlm."
Option Explicit

Private Sub B_SlmbAli()
'GoSub T1
'GoSub T2
'GoSub T3
GoSub T4
Exit Sub
Dim Slmb$()
T1:
    Slmb = Sy( _
        "Function A(AA): A = 1:End function", _
        "Sub SlmzFfn(Ffn): VVSlmBlkFfn Ffn: End Sub", _
        "Sub SlmM():                 SlmzM Cmd:                                      End Sub", _
        "Sub SlmzM(M As CodeModule): RplMdLyopt M, OptSrcAliSlm(Src(M)), IsPjSav:=True: End Sub", _
        "Sub SlmzMdn(Mdn$):          SlmzM Md(Mdn):                        End Sub")
    Ept = Sy( _
        "Function A(AA):             A = 1:                                                 End function", _
        "Sub SlmzFfn(Ffn):               VVSlmBlkFfn Ffn:                                   End Sub", _
        "Sub SlmM():                     SlmzM Cmd:                                         End Sub", _
        "Sub SlmzM(M As CodeModule):     RplMdLyopt M, OptSrcAliSlm(Src(M)), IsPjSav:=True: End Sub", _
        "Sub SlmzMdn(Mdn$):              SlmzM Md(Mdn):                                     End Sub")
    GoTo Tst
T2:
    Slmb = Sy("Sub AA:              End Sub")
    Ept = Sy("Sub AA: End Sub")
    GoTo Tst
T3:
    Slmb = Sy("Function AA(): AA = 1:             End Function")
    Ept = Sy("Function AA(): AA = 1: End Function")
    GoTo Tst
T4:
    Slmb = Sy( _
        "Function Acs() As Access.Application: Set Acs = Access.Application: End Function" & _
        "Function AcsDb(Db As Database) As Access.Application: Set AcsDb = AcsFb(Db.Name): End Function")
    Ept = Sy( _
        "Function Acs() As Access.Application:                   Set Acs = Access.Application: End Function" & _
        "Function AcsDb(Db As Database) As Access.Application: Set AcsDb = AcsFb(Db.Name):     End Function")
    GoTo Tst
Tst:
    Act = SlmbAli(Slmb)
    C
    Return
End Sub

Private Sub B_WDr4Slmln()
Dim Slmln$
GoSub T1
GoSub T2
GoSub T3
GoSub T4
GoSub T5
Exit Sub
T1:
    Slmln = "Private Sub Cmd_Opn_Fxw_Click(): MaxvFx xWFx:  End Sub"
    Ept = Sy( _
        "Private Sub Cmd_Opn_Fxw_Click(): ", _
        "", _
        "MaxvFx xWFx: ", _
        "End Sub")
    GoTo Tst
T2:
    Slmln = "Private Sub Cmd_Opn_PthOup_Click():    End Sub"
    Ept = Sy( _
        "Private Sub Cmd_Opn_PthOup_Click(): ", _
        "", _
        "", _
        "End Sub")
    GoTo Tst
T3:
    Slmln = "Private Function AA$():    AA = AA:           End Function"
    Ept = Sy( _
        "Private Function AA$(): ", _
        "AA = ", _
        "AA: ", _
        "End Function")
    GoTo Tst
T4:
    Slmln = "Private Function sampYmd() As Ymd: SampYmd = Ymd(19, 12, 24): End Function"
    Ept = Sy( _
        "Private Function sampYmd() As Ymd: ", _
        "SampYmd = ", _
        "Ymd(19, 12, 24): ", _
        "End Function")
    GoTo Tst
T5:
    Slmln = "Function Acs() As Access.Application:                 Set Acs = Access.Application: End Function"
    Ept = Sy( _
        "Function Acs() As Access.Application: ", _
        "Set Acs = ", _
        "Access.Application: ", _
        "End Function")
Tst:
    Act = WDr4Slmln(Slmln)
    D Act
    C
    Return
End Sub
Private Function WDr4Slmln(Slmln) As String() ' Always return 4 elements
Dim L$: L = Slmln
Dim K$: K = MthkdL(L)
Dim IsFun As Boolean: IsFun = K = "Function"
Dim EMthColon$, ELhs$, EBefEnd$
Dim M$
EMthColon = ShfBef(L, ":", InlSep:=True): If EMthColon = "" Then M = "*EMthColon must contain value": GoTo M
EMthColon = EMthColon & " "
     ELhs = WShfLHS(L, IsFun):                 If ELhs <> "" Then ELhs = ELhs & " "
  EBefEnd = Trim(ShfBef(L, "End " & K)):    If EBefEnd <> "" Then EBefEnd = EBefEnd & " "
                                                   If L = "" Then M = "*Rest must contain value": GoTo M
WDr4Slmln = Sy(EMthColon, ELhs, EBefEnd, L)
Exit Function
M: Inf CSub, M, "Slmln", Slmln: Stop
RaisePgmEr
End Function
Private Function WShfLHS$(OLn$, IsFun As Boolean)
If IsFun Then WShfLHS = LTrim(ShfBef(OLn, "=", InlSep:=True))
End Function

Function SlmbAli(Slmb$()) As String()
Dim Dy()
    Dim L: For Each L In Itr(Slmb)
        PushI Dy, WDr4Slmln(L)
    Next
SlmbAli = FmtLndy(Dy, CiiAliR:=1, NoIx:=True, Sep:="")
If Si(SlmbAli) <> Si(Slmb) Then Stop
End Function

Function SlmboptAli(Slmb$()) As Lyopt: SlmboptAli = LyoptOldNew(Slmb, SlmbAli(Slmb)): End Function


