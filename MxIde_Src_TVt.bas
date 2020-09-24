Attribute VB_Name = "MxIde_Src_TVt"
Option Compare Text
Option Explicit

Const CMod$ = "MxIde_Src_TVtTy."
Type TVt: Tyc As String: Tyn As String: IsAy As Boolean: End Type  ' Deriving(Ay) #Variable
Function TVtTyBkt(Tyc$, TyBkt$) As TVt
If HasSfx(TyBkt, "()") Then
    TVtTyBkt = TVt(Tyc, RmvSfx(TyBkt, "()"), True)
Else
    TVtTyBkt = TVt(Tyc, TyBkt, False)
End If
End Function
Function TVt(Tyc, Tyn, IsAy) As TVt
With TVt
    .Tyc = Tyc
    .IsAy = IsAy
    .Tyn = Tyn
End With
End Function
Function SiTVt&(A() As TVt): On Error Resume Next: SiTVt = UBound(A) + 1: End Function
Function UbTVt&(A() As TVt): UbTVt = SiTVt(A) - 1: End Function

Function ShtVsfx$(Vsfx$)
If Vsfx = "" Then Exit Function
Dim L$: L = Vsfx
Select Case True
Case L = " As Boolean":: ShtVsfx = "^"
Case L = " As Boolean()": ShtVsfx = "^()"
Case Else
    IsShfPfx L, " As "
    ShtVsfx = L
End Select
End Function

Function ShtVsfxTVt$(T As TVt)
With T
    Dim B$: B = BktIf(.IsAy)
    Dim O$
    Select Case True
    Case .Tyc = "" And .Tyn = "": O = ":V" & B
    Case .Tyc = "": O = ":" & .Tyn & B
    Case Else: O = .Tyc & B
    End Select
End With
ShtVsfxTVt = O
End Function
Function TVtVsfx(Vsfx$) As TVt
Const CSub$ = CMod & "TVtVsfx"
Dim S$: S = Vsfx
Dim Tyc$: Tyc = ShfTyc(S): If S = "" Then GoTo X
Dim IsAy As Boolean: IsAy = IsShfBkt(S): If IsAy Then If S = "" Then GoTo X
If Not IsShfAs(S) Then Thw CSub, "Given @Vsfx is invalid: after  shifting Varn, Tyc, (), there is some left, but is not ()", "Vsfx", Vsfx
Dim IsNew As Boolean: IsNew = IsShfPfx(S, "New ")
Dim Tyn$: Tyn = ShfNmDot(S)
Select Case True
Case S = "()": If IsAy Then Thw CSub, "There is 2 () in Vsfx", "Vsfx", Vsfx
           IsAy = True
Case S = "":
Case Else
    If IsShfPfx(S, " * ") Then
        If Not IsNm(S) Then
            If Not IsNumeric(S) Then Thw CSub, "Vsfx has [String * ..], but [..] is not numeric nor name(assume the name is a constant)", "Vsfx", Vsfx
        End If
    Else
        Thw CSub, "After shifting As-Nm, there is still something left", "Vsfx", Vsfx
    End If
End Select
X: With TVtVsfx
    .IsAy = IsAy
    .Tyc = Tyc
    .Tyn = Tyn
End With
End Function
