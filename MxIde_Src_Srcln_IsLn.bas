Attribute VB_Name = "MxIde_Src_Srcln_IsLn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Srcln_IsLn."
Function IsLnPrp(L) As Boolean
IsLnPrp = MthkdL(L) = "Property"
End Function

Private Sub B_IsLnMth()
GoTo Z
Dim A$
A = "Function IsLnMth(A) As Boolean"
Ept = True
GoSub Tst
Exit Sub
Tst:
    Act = IsLnMth(A)
    C
    Return
Z:
    Dim L, O$()
    For Each L In SrcMC
        If IsLnMth(CStr(L)) Then
            PushI O, L
        End If
    Next
    Brw O
    Return
End Sub

Function IsLnImp(L) As Boolean: IsLnImp = HasPfx(L, "Implements "): End Function
Function IsLnOpt(L) As Boolean
If Not HasPfx(L, "Option ") Then Exit Function
Select Case True
Case _
    HasPfx(L, "Option Explicit"), _
    HasPfx(L, "Option Compare Text"), _
    HasPfx(L, "Option Compare Binary"), _
    HasPfx(L, "Option Compare Database")
    IsLnOpt = True
End Select
End Function

Function IsLnMthPub(L) As Boolean
Dim Ln$: Ln = L
Dim Mdy$: Mdy = ShfMdy(Ln): If Mdy <> "" And Mdy <> "Public" Then Exit Function
IsLnMthPub = TakMthkd(Ln) <> ""
End Function

Function IsLnMth(L) As Boolean:          IsLnMth = MthTyLn(L) <> "":           End Function
Function IsLnMthn(L, Mthn_) As Boolean: IsLnMthn = MthnL(L) = Mthn_:           End Function
Function IsLnFun(L) As Boolean:          IsLnFun = MthTyLn(L) = "Function":    End Function
Function IsLnUdt(L) As Boolean:          IsLnUdt = HasPfx(RmvMdy(L), "Type "): End Function

Function IsLnNonSrc(L) As Boolean
IsLnNonSrc = True
If HasPfx(L, "Option ") Then Exit Function
Dim Ln$: Ln = Trim(L)
If Ln = "" Then Exit Function
IsLnNonSrc = False
End Function

Function IsLnVmk(L) As Boolean:   IsLnVmk = ChrFst(LTrim(L)) = "'": End Function
Function IsLnBlnk(L) As Boolean: IsLnBlnk = Trim(L) = "":           End Function
Function IsLnBlnkOrVmk(L) As Boolean
Dim Ln$: Ln = LTrim(L)
Select Case True
Case Ln = "", ChrFst(Ln) = "'": IsLnBlnkOrVmk = True
End Select
End Function
Function IsLnVmkOrBlnk(Ln) As Boolean
Select Case True
Case IsLnBlnk(Ln), IsLnVmk(Ln): IsLnVmkOrBlnk = True: Exit Function
End Select
End Function
Function IsLnCd(L) As Boolean: IsLnCd = Not IsLnVmkOrBlnk(L): End Function

Function IsLnNonOpt(Ln) As Boolean
If Not IsLnCd(Ln) Then Exit Function
If HasPfx(Ln, "Option") Then Exit Function
IsLnNonOpt = True
End Function
