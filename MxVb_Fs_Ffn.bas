Attribute VB_Name = "MxVb_Fs_Ffn"
Option Compare Text
Option Explicit
Public Fso As New FileSystemObject
Const CMod$ = "MxVb_Fs_Ffn."
Enum eFilCpr: eFilCprByt: eFilCprTimSi: End Enum
Public Const SepPth$ = "\"
Function FfnPthFn$(Pth, Fn): FfnPthFn = PthEnsSfx(Pth) & Fn: End Function
Function FfnChkExist$(Ffn, Optional Fun$, Optional Kd$)
ChkFfnExi Ffn, Fun, Kd
FfnChkExist = Ffn
End Function
Function EryFfnMis(Ffn, Optional Kd$ = "File") As String()
If HasFfn(Ffn) Then Exit Function
Erase XX
X FmtQQ("? not found", Kd)
X vbTab & "Path : " & Pth(Ffn)
X vbTab & "File : " & Fn(Ffn)
EryFfnMis = XX
Erase XX
End Function

Function FfnRplExt$(Ffn, NewExt): FfnRplExt = Ffnn(Ffn) & NewExt: End Function

Function IsExtInAp(Ffn, ParamArray Ap()) As Boolean: Dim Av(): Av = Ap: IsExtInAp = IsInAv(Ext(Ffn), Av): End Function
Function IsInAv(V, Av()) As Boolean: IsInAv = HasEle(Av, V): End Function

Function IsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
IsInAp = HasEle(Av, V)
End Function

Function PthUp$(Pth, Optional NUp% = 1)
Dim O$: O = Pth
Dim J%
For J = 1 To NUp
    O = PthPar(O)
Next
PthUp = O
End Function

Function PthFfnPar$(Ffn): PthFfnPar = PthPar(Pth(Ffn)): End Function

Sub ChkIsFxa(Ffn, Optional Fun$)
If Not IsFxa(Ffn) Then Thw Fun, "Given Ffn is not Fxa", "Ffn", Ffn
End Sub
Function FxyFfny(Ffny$()) As String()
Dim Ffn: For Each Ffn In Itr(Ffny)
    If IsFx(Ffn) Then PushI FxyFfny, Ffn
Next
End Function
Function FbyFfny(Ffny$()) As String()
Dim Ffn: For Each Ffn In Itr(Ffny)
    If IsFb(Ffn) Then PushI FbyFfny, Ffn
Next
End Function

Function StrSiDotDTim$(Ffn)
If HasFfn(Ffn) Then StrSiDotDTim = StrDte(FfnTim(Ffn)) & "." & FfnLen(Ffn)
End Function

Function HasSsExt(Ffn, SsExt$) As Boolean: HasSsExt = HasEleStr(SySs(SsExt), Ext(Ffn), eCasIgn): End Function

Sub OvrWrt(Ffn$, ShdOvrWrt As Boolean)
If ShdOvrWrt Then
    DltFfnIf Ffn
Else
    ChkNoFfn Ffn
End If
End Sub

Function FfnUp$(Ffn): FfnUp = PthPar(Pth(Ffn)) & Fn(Ffn): End Function
