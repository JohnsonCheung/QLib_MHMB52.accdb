Attribute VB_Name = "MxIde_Mthn_shttyMthn2"
Option Compare Text
Option Explicit
Function Mi2TyNmPubL$(L)
Dim S$: S = L
If Not IsShfPub(S) Then Exit Function
Dim ShtTy$: ShtTy = ShfShtMthTy(S): If ShtTy = "" Then Exit Function
Dim N$: N = TakNm(S)
    Select Case Left(N, 2)
    Case "A_", "B_": Exit Function
    End Select
    If Right(N, 5) = "__Tst" Then Exit Function
Mi2TyNmPubL = ShtTy & " " & N
End Function
Private Sub Mi2yTyNmPubPC__Tst(): Vc Mi2yTyNmPubPC: End Sub
Function Mi2yTyNmPubPC() As String()
Dim L: For Each L In SrcPC
    PushNB Mi2yTyNmPubPC, Mi2TyNmPubL(L)
Next
End Function
