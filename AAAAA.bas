Attribute VB_Name = "AAAAA"
'#Cii:Column-Integer-Index-in-Ss-Format:String#
'C=Column i=Integer ii=dbl-letter-at-end-means-in-format-of-Ss
Option Compare Text
Option Explicit
Const CMod$ = "AA."

Function DwColNoSng(D As Drs, C$) As Drs
'@A : has a column-C
'Ret   : sam stru as A and som row removed.  rmv row are its col C value is Single. @@
Dim Col(): Col = DcDrs(D, C)
Dim Sng(): Sng = AwSng(Col)
DwColNoSng = DeInVy(D, C, Sng)
End Function
Function DwFstDrWhCeq(A As Drs, C$, Eqval) As Variant()
Dim Cix%: Cix = IxEle(A.Fny, C)
Dim Rix&: Rix = RixWhColEq(A.Dy, Cix, Eqval)
DwFstDrWhCeq = A.Dy(Rix)
End Function

Function RecFstWhDrsCeq(A As Drs, C$, Eqval) As Rec
RecFstWhDrsCeq = Rec(A.Fny, DwFstDrWhCeq(A, C, Eqval))
End Function


Function AvAy(Ay) As Variant()
If IsAv(Ay) Then AvAy = Ay: Exit Function
Dim I: For Each I In Itr(Ay)
    Push AvAy, I
Next
End Function
