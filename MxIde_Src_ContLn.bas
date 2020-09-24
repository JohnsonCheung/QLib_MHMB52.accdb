Attribute VB_Name = "MxIde_Src_ContLn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_ContLn."
Function ContlnyPC() As String(): ContlnyPC = Contlny(SrcPC): End Function
Sub DltContln(M As CodeModule, Lno&)
M.DeleteLines Lno, NContln(SrcM(M), Lno)
End Sub

Function ContlnIx$(Src$(), Ix)
If Ix = -1 Then Exit Function
Dim N%: N = NContln(Src, Ix)
If N = 1 Then
    ContlnIx = Src(Ix)
    Exit Function
End If
Dim O$()
    Dim J&: For J = Ix To Ix + N - 2
        PushI O, RmvLas(Src(J))
    Next
    PushI O, Src(Ix + N - 1)
ContlnIx = Jn(O)
End Function

Private Sub B_Contlny(): Vc Contlny(SrcPC): End Sub
Function Contlny(Src$()) As String()
Dim IsContPrv As Boolean, IsContCur As Boolean
Dim O$(), IxO&
    IxO = -1
    Dim L: For Each L In Itr(Src)
        IsContCur = ChrLas(L) = "_"
        If IsContCur Then L = RmvLas(L)
        If IsContPrv Then
            O(IxO) = O(IxO) & LTrim(L)
        Else
            PushI O, L
            IxO = IxO + 1
        End If
        IsContPrv = IsContCur
    Next
Contlny = O
End Function

Function ContlnM$(M As CodeModule, Lno&)
'123456789012345678901234567890123456789012345678901234567890 _
123456789012345678901234567890123456789012345678901234567890 _
123456789012345678901234567890123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
123456789012345678901234567890 _
sdfsdf
Const CSub$ = CMod & "ContlnM"
If Lno = 0 Then Exit Function
Dim O$
Dim J&: For J = Lno To M.CountOfLines
    Dim L$: L = M.Lines(J, 1)
    If ChrLas(L) <> "_" Then
        If O = "" Then
            ContlnM = L
        Else
            ContlnM = O & LTrim(L)
        End If
        Exit Function
    End If
    O = O & RmvLas(LTrim(L))
Next
Thw CSub, "Las Ln of @Md has [_] at end", "@Md", Mdn(M)
End Function

Function IxSrclnNxt&(Src$(), Optional Ix = 0): IxSrclnNxt = NContln(Src, Ix): End Function
Private Sub B_IxCdlnNxt()
GoSub T1
Exit Sub
Dim Src$(), Ix
T1:
    Src = Sy( _
        "AAA _", _
        "AA", _
        "' sdlf", _
        "    'sdlkfj", _
        "  A")
    Ix = 0
    Ept = 4
    GoTo Tst
Tst:
    Act = IxCdlnNxt(Src, Ix)
    Debug.Assert Act = Ept
    Return
End Sub
Private Sub B_IxSrclnNxt()
GoSub T1
Exit Sub
Dim Src$(), Ix
T1:
    Src = Sy( _
        "AAA _", _
        "AA", _
        "' sdlf", _
        "    'sdlkfj", _
        "  A")
    Ix = 0
    Ept = 2
    GoTo Tst
Tst:
    Act = IxSrclnNxt(Src, Ix)
    Debug.Assert Act = Ept
    Return
End Sub
Function IxCdlnNxt&(Src$(), Optional Ix = 0)
Dim O&: O = IxSrclnNxt(Src, Ix)
For O = O To UB(Src)
    If Not IsLnVmk(Src(O)) Then IxCdlnNxt = O: Exit Function
Next
Thw CSub, "Given Src does not have nxt-cdl", "Src Bix", Src, Ix
End Function

Function NContlnLno%(M As CodeModule, Lno)
Const CSub$ = CMod & "NContlnLno"
Dim J&, O%: For J = Lno To M.CountOfLines
    O = O + 1
    If ChrLas(M.Lines(J, 1)) <> "_" Then
        NContlnLno = O
        Exit Function
    End If
Next
Thw CSub, "las ln of Md cannot have [_] at end", "Las-Ele-Of-Md Md", M.Lines(M.CountOfLines, 1), Mdn(M)
End Function

Function NContln(Src$(), Ix) As Byte
Const CSub$ = CMod & "NContln"
Dim J&, O&: For J = Ix To UB(Src)
    O = O + 1
    If ChrLas(Src(J)) <> "_" Then
        NContln = O
        Exit Function
    End If
Next
Thw CSub, "las ele of Src cannot have [_] at end", "Src-LasELe", EleLas(Src)
End Function

