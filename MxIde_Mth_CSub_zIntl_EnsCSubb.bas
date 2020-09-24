Attribute VB_Name = "MxIde_Mth_CSub_zIntl_EnsCSubb"
Option Compare Database
Option Explicit

Private Function CSubln$(Mthn$)
CSubln = FmtQQ("Const CSub$ = CMod & ""?""", Mthn)
End Function
Private Function IxInsCSub%(Mthy$()): IxInsCSub = IxCdlnNxt(Mthy): End Function

Function CSubbEns(Mthy$()) As String()
Const CSub$ = CMod & "CSubbEns"
If Si(Mthy) = 1 Then CSubbEns = Mthy: Exit Function
Dim Ix%: Ix = IxCsntn(Mthy, "CSub")
If IsUseCSub(Mthy) Then
    Dim L$: L = CSubln(MthnL(Mthy(0)))
    If Ix = -1 Then
        Dim IxIns%: IxIns = IxInsCSub(Mthy)
        CSubbEns = AyIns(Mthy, L, IxIns)
    Else
        CSubbEns = SySetStr(Mthy, Ix, L)
    End If
Else
    If Ix > 0 Then
        CSubbEns = AeIx(Mthy, Ix)
    Else
        CSubbEns = Mthy
    End If
End If
End Function


Private Sub B_CSubbEns()
Const CSub$ = CMod & "B_CSubbEns"
GoSub T1
'GoSub ZZ
Exit Sub
Dim Mthy$()
ZZ:
    Dim Lyy(): Lyy = LyyMthPC
    Dim O$()
    Dim I: For Each I In Lyy
        Dim MthI$(): MthI = I
        Dim MthlNew$: MthlNew = JnCrLf(CSubbEns(MthI))
        Dim MthlOld$: MthlOld = JnCrLf(MthI)
        If MthlOld <> MthlNew Then
            PushI O, LinesUL(MthlOld) & vbCrLf & MthlNew
        End If
    Next
    Stop
    BrwLsy O
    Return
T1:
    Erase Mthy
    PushI Mthy, "Function FxWMisEr(Fx, W, Optional FxKd$ = ""Excel file"") As String()"
    PushI Mthy, "Dim W: For Each W In Tml(W)"
    PushI Mthy, "    PushIAy FxWMisEr, FxwMisEr(Fx, W, FxKd)"
    PushI Mthy, "Next"
    PushI Mthy, "Thw CSub, ""AA"""
    PushI Mthy, "End Function"
    GoTo Tst
Tst:
    Act = CSubbEns(Mthy)
    D Act
    Stop
    C
    Return
End Sub

Function IsUseCSub(Mthy$()) As Boolean
Dim L: For Each L In Itr(Mthy)
    If IsLnUseCSub(L) Then IsUseCSub = True: Exit Function
Next
End Function

Private Sub B_IsLnUseCSub()
Dim LyUse$(), LyNot$()
    PushIAy LyNot, LyUL("#1 NotUse CSub", "=")
    PushIAy LyUse, LyUL("#2 IsUse CSub", "=")
    Dim C As VBComponent: For Each C In CPj.VBComponents
        DoEvents
        Dim Src$(): Src = SrcCmp(C)
        Dim N$: N = C.Name
        Dim J%: For J = 0 To UB(Src)
            Dim L$: L = Src(J)
            Dim P As P12: P = P12Ssub(L, "CSub")
            Dim Ly2$()
            If P.P1 > 0 Then
                If IsLnUseCSub(L) Then
                    P = P12Ssuby(L, SsubyCSub)
                    Ly2 = Jrcy(N, J + 1, P, L)
                    PushIAy LyUse, Ly2
                Else
                    If Not HasCnstnL(L, "CSub") And Not IsLnVmk(L) Then        'Don't put [COnst CSub$=] to LyNot
                        Ly2 = Jrcy(N, J + 1, P, L)
                        PushIAy LyNot, Ly2
                    End If
                End If
            End If
        Next
    Next
Dim Ly$()
    Ly = SyAdd(LyNot, LyUse)
    PushI Ly, "Count =" & Si(Ly) / 2
    PushI Ly, "IsUse =" & Si(LyUse) / 2
    PushI Ly, "NotUse=" & Si(LyNot) / 2
    
VcAy Ly, "IsLnUseCSub__Tst "
End Sub

Function CSubboptEns(Mthy$()) As Lyopt: CSubboptEns = LyoptOldNew(Mthy, CSubbEns(Mthy)): End Function
Private Function SsubyCSub() As String()
Static X As Boolean, Y ' Y cannot dcl as Y$(), otherwise, next time Y will be empyt!!
If Not X Then X = True: Y = Tmy("[CSub, ] [, CSub] (CSub [ThwImposs CSub]")
SsubyCSub = Y
End Function
Private Function IsLnUseCSub(L) As Boolean
If IsLnVmk(L) Then Exit Function
IsLnUseCSub = HasSsubyOr(L, SsubyCSub)
End Function
Function IsLnCSub(L) As Boolean: IsLnCSub = HasCnstnL(L, "CSub"): End Function

