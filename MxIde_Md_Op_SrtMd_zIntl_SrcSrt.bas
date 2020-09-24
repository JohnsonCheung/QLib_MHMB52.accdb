Attribute VB_Name = "MxIde_Md_Op_SrtMd_zIntl_SrcSrt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_SrcSrt."

Function SrcSrt(Src$()) As String()
Dim Di As Dictionary: Set Di = DiMthlSrc(Src)
Dim DiSrt As Dictionary: Set DiSrt = DiSrtStr(Di)
Dim Mthlsy$(): Mthlsy = DivyStr(DiSrt)
SrcSrt = LyLsy(WMthlsyInsBlnkEle(Mthlsy))
End Function
Private Function WMthlsyInsBlnkEle(Mthlsy$()) As String()
If Si(Mthlsy) = 0 Then Exit Function
PushI WMthlsyInsBlnkEle, Mthlsy(0)
Dim J%: For J = 1 To UB(Mthlsy)
    If IsLinesMoreThan1Ln(Mthlsy(J - 1)) Then
        GoSub PushBlnkLn
    ElseIf IsLinesMoreThan1Ln(Mthlsy(J)) Then
        GoSub PushBlnkLn
    End If
    PushI WMthlsyInsBlnkEle, Mthlsy(J)
Next
Exit Function
PushBlnkLn: PushI WMthlsyInsBlnkEle, "": Return
End Function
Function SrclSrt$(Src$()): SrclSrt = JnCrLf(SrcSrt(Src)): End Function
Private Sub B_SrcSrt()
GoSub ZZ
Exit Sub
ZZ_Dcl_Bef_And_Aft_Srt:
    Const Mdn$ = "DqStrRe"
    Dim SrcBef$() ' Src
    Dim SrcAft$() ' Src->Srt
    Dim DclBef$() 'Src->Dcl
    Dim DclAft$() 'Src->Src->Dcl
    SrcBef = SrcMC
    SrcAft = SrcSrt(SrcBef)
    DclBef = DclSrc(SrcBef)
    DclAft = DclSrc(SrcAft)
    Ass JnCrLf(DclBef) = JnCrLf(DclAft)
    Return
ZZ:
    Dim C As VBComponent: For Each C In CPj.VBComponents
        Dim Src$(), Cmpn$
        Src = SrcCmp(C)
        Cmpn = C.Name
        GoSub ZZ_Tst
    Next
    Return
ZZ_Tst:
    SrcAft = SrcSrt(Src)
    If JnCrLf(Src) = JnCrLf(SrcAft) Then
        Debug.Print Cmpn, "Is Same of before and after sorting ......"
        Return
    End If
    If Si(SrcAft) <> 0 Then
        If EleLas(SrcAft) = "" Then
            Dim Pfx
            Pfx = Array("There is non-blank-line at end after sorting", "Md=[" & Cmpn & "=====")
            BrwAy AyAddAp(Pfx, SrcAft)
            Stop
        End If
    End If
    Dim A$(), B$(), Ii
    A = AyMinus(Src, SrcAft)
    B = AyMinus(SrcAft, Src)
    Debug.Print
    If Si(A) = 0 And Si(B) = 0 Then Return
    If Si(AeEmpEle(A)) <> 0 Then
        Debug.Print "Si(A)=" & Si(A)
        BrwAy A
        Stop
    End If
    If Si(AeEmpEle(B)) <> 0 Then
        Debug.Print "Si(B)=" & Si(B)
        BrwAy B
        Stop
    End If
    Return
End Sub

Function SrcloptSrtM(M As CodeModule) As Stropt
Dim S$(): S = SrcM(M)
Dim Newl$: Newl = JnCrLf(SrcSrt(S))
Dim Oldl$: Oldl = JnCrLf(S)
SrcloptSrtM = StroptOldNew(Oldl, Newl)
End Function
Private Function SrcSrtM(M As CodeModule) As String():   SrcSrtM = SrcSrt(SrcM(M)):     End Function
Function SrcSrtMC() As String():                        SrcSrtMC = SrcSrtM(CMd):        End Function
Function SrcSrtMdn(Mdn$) As String():                  SrcSrtMdn = SrcSrtM(MdMdn(Mdn)): End Function
Sub VcSrcSrtMC():                                                  VcAy SrcSrtMC:       End Sub
