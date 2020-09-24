Attribute VB_Name = "MxDta_Da_Dy_DySrt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dy_DySrt."
Function DySrtCii(Dy(), TmlCiiHyp$) As Variant()
If TmlCiiHyp = "" Then DySrtCii = Dy: Exit Function
Dim K() As Srkey: K = WSrkeyyCiiHyp(TmlCiiHyp)
DySrtCii = DySrtKeyy(Dy, K)
End Function
Private Function WSrkeyyCiiHyp(TmlCiiHyp$) As Srkey()
Dim A$(): A = SplitSpc(TmlCiiHyp)
Dim O() As Srkey
    ReDim O(UB(A))
    Dim J%: For J = 0 To UB(A)
        O(J) = WSrkeyCiHyp(A(J))
    Next
WSrkeyyCiiHyp = O
End Function
Private Function WSrkeyCiHyp(CiHyp) As Srkey
With WSrkeyCiHyp
    If HasSfx(CiHyp, "-") Then
        .Ci = RmvLas(CiHyp)
        .IsDes = True
    Else
        .Ci = CiHyp
    End If
End With
End Function

Function DySrtKeyy(Dy(), K() As Srkey) As Variant()
Dim D(): D = DySel(Dy, CiySrkey(K))
Dim Rxy&(): Rxy = RxySrtDy(D, WIsDesyKeyy(K))
DySrtKeyy = AwIxy(Dy, Rxy)
End Function
Private Function WIsDesyKeyy(K() As Srkey) As Boolean() ' ret same ele count as @U
Dim O() As Boolean
Dim U%: U = UbSrkey(K): ReDim O(U)
Dim J%: For J = 0 To U
    With K(J)
        If .IsDes Then
            O(.Ci) = True
        End If
    End With
Next
WIsDesyKeyy = O
End Function

Function DySrtSngDc(Dy(), Ci, Optional IsDes As Boolean) As Variant()
Dim K() As Srkey: K = WSrkeyySng(Ci, IsDes)
DySrtSngDc = DySrtKeyy(Dy, K)
End Function
Function DySrt(Dy(), Optional IsDes As Boolean) As Variant()
Dim K() As Srkey: K = WSrkeyyDyIsDes(Dy, IsDes)
DySrt = DySrtKeyy(Dy, K)
End Function
Private Function WSrkeyySng(Ci, IsDes) As Srkey(): PushSrkey WSrkeyySng, Srkey(Ci, IsDes): End Function
Private Function WSrkeyyDyIsDes(Dy(), IsDes As Boolean) As Srkey()
If Si(Dy) = 0 Then Exit Function
Dim O() As Srkey
    Dim UDc%: UDc = UDcDy(Dy): If UDc = -1 Then Exit Function
    ReDim O(UDc)
    
    If IsDes Then
        Dim J%: For J = 0 To UDc
            With O(J)
                .IsDes = True
            End With
        Next
    End If
WSrkeyyDyIsDes = O
End Function
