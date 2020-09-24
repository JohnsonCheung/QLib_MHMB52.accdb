Attribute VB_Name = "MxVb_Dta_Bei"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Bei."
Type Bei: Bix As Long: Eix As Long: End Type
Function SiBei&(A() As Bei): On Error Resume Next: SiBei = UBound(A) + 1: End Function
Function UbBei&(A() As Bei): UbBei = SiBei(A) - 1: End Function

Function RepBei$(A As Bei): RepBei = FmtQQ("Bei ? ?", A.Bix, A.Eix): End Function
Function RepyBeiy$(A() As Bei)
Dim O$()
Dim J&: For J = 0 To UbBei(A)
    With A(J)
        PushI O, FmtQQ("?, ?", .Bix, .Eix)
    End With
Next
RepyBeiy = FmtQQ("BeiAy(?)", JnCmaSpc(O))
End Function
Function BeiyBooly(Booly() As Boolean) As Bei()
Dim U&: U = UB(Booly): If U = -1 Then Exit Function
Dim B&(): B = W2BoolyBixy(Booly)
Dim J&: For J = 0 To UB(B)
    PushBei BeiyBooly, Bei(B(J), W2Eix(Booly, B(J)))
Next
End Function
Private Function W2BoolyBixy(Booly() As Boolean) As Long()
Dim Las As Boolean, Fst As Boolean
Dim J&: For J = 0 To UB(Booly)
    Select Case True
    Case (Fst Or Not Las) And Booly(J): Fst = False: PushI W2BoolyBixy, J
    End Select
    If Fst Then Fst = False
    Las = Booly(J)
Next
End Function
Private Function W2Eix&(Booly() As Boolean, Bix&)
Dim U&: U = UB(Booly)
Dim J&: For J = Bix + 1 To U
    If Not Booly(J) Then W2Eix = J - 1: Exit Function
Next
W2Eix = U
End Function
Private Function W1BixPfxy&(Ly, Pfxy$())
Dim J&: For J = 0 To UB(Ly)
    If HasPfxySpc(Ly(J), Pfxy) Then W1BixPfxy = J: Exit Function
Next
W1BixPfxy = -1
End Function
Private Function W1EixPfxy&(Ly, Bix%, Pfxy$())
If Bix = -1 Then W1EixPfxy = -1: Exit Function
Dim U&: U = UB(Ly)
Dim J&: For J = Bix + 1 To U
    If Not HasPfxySpc(Ly(J), Pfxy) Then W1EixPfxy = J - 1: Exit Function
Next
W1EixPfxy = U
End Function
Function BeiEmp() As Bei
BeiEmp.Bix = -1
BeiEmp.Eix = -1
End Function

Function Bei(Bix, Eix) As Bei
Select Case True
Case 0 > Bix, -1 > Eix, Bix > Eix: Thw CSub, "Bix and Eix must >0 and Bix must <= Eix", "Bix Eix", Bix, Eix
End Select
Bei.Bix = Bix
Bei.Eix = Eix
End Function
Sub PushBei(O() As Bei, M As Bei)
Dim N&: N = SiBei(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Sub PushBeiy(O() As Bei, A() As Bei)
Dim J&: For J = 0 To UbBei(A)
    PushBei O, A(J)
Next
End Sub
Function BeiyAdd(A As Bei, B As Bei) As Bei()
PushBei BeiyAdd, A
PushBei BeiyAdd, B
End Function

Function CntBei&(A As Bei)
Const CSub$ = CMod & "CntBei"
CntBei = A.Eix - A.Bix
If CntBei < 0 Then Thw CSub, "Given @Bei has negative Cnt", "Bix Eix", A.Bix, A.Eix
End Function
Function IsEqBeiy(A() As Bei, B() As Bei) As Boolean
Dim U&: U = UbBei(A)
If U <> UbBei(B) Then Exit Function
Dim J&: For J = 0 To U
    If Not IsEqBei(A(J), B(J)) Then Exit Function
Next
IsEqBeiy = True
End Function
Function IsEmpBei(A As Bei) As Boolean: IsEmpBei = A.Bix = -1 And A.Eix = -1: End Function
Function IsBeiyInOrd(A() As Bei) As Boolean
Dim J&: For J = 0 To UbBei(A)
    With LcntBei(A(J))
        If .Lno = 0 Then Exit Function
        If .Cnt = 0 Then Exit Function
        If .Lno + .Cnt > LcntBei(A(J + 1)).Lno Then Exit Function
    End With
Next
IsBeiyInOrd = True
End Function

Function IsEqBei(A As Bei, B As Bei) As Boolean
With A
    If .Bix <> B.Bix Then Exit Function
    If .Eix <> B.Eix Then Exit Function
End With
IsEqBei = True
End Function

Sub BrwBeiy(A() As Bei): BrwAy RepyBeiy(A): End Sub
