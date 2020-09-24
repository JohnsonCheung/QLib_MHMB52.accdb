Attribute VB_Name = "MxIde_Src_TyDfn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_TyDfn."

Function IsLnTyDfn(L) As Boolean: IsLnTyDfn = TyDfnn(L) <> "": End Function
Function TyDfnnyPC() As String(): TyDfnnyPC = TyDfnny(SrcPC):  End Function

Function TyDfnny(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB TyDfnny, TyDfnn(L)
Next
End Function
Function TyDfnn$(L)
If Left2(L) <> "':" Then Exit Function
Dim T$: T = Tm1(L): If Len(T) = 2 Then Exit Function
If ChrLas(T) <> ":" Then Exit Function
TyDfnn = RmvFst(T)
End Function

Function IsLnTyDfnRmk(Ln) As Boolean
If ChrFst(Ln) <> "'" Then Exit Function
If ChrFst(LTrim(RmvFst(Ln))) <> "!" Then Exit Function
IsLnTyDfnRmk = True
End Function

Function IsTyDfnn(Nm$) As Boolean
':TyDfnn: :Nm  #TyDfn-Name# It must be from a str with fst2chr is [':], and then non-space-chr, and then [:].
'         Then non-space char is :TyDfnNm
Select Case True
Case Left2(Nm) <> "':", ChrLas(Nm) <> ":"
Case Else: IsTyDfnn = True
End Select
End Function

Function IsMemn(Tm$) As Boolean
If Len(Tm) > 3 Then
    If ChrFst(Tm) = "#" Then
        If ChrLas(Tm) = "#" Then
            IsMemn = True
        End If
    End If
End If
End Function
