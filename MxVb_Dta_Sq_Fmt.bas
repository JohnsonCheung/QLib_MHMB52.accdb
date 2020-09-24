Attribute VB_Name = "MxVb_Dta_Sq_Fmt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Sq_Fmt."

Sub BrwSq(Sq()):                                                  Brw FmtSq(Sq):            End Sub
Function FmtSq(Sq(), Optional SepChr$ = " ") As String(): FmtSq = WL£y(WSqAli(Sq), SepChr): End Function
Private Function WL£y(Sq(), SepChr$) As String()
Dim NC&: NC = UBound(Sq, 2)
Dim R&: For R = 1 To UBound(Sq, 1)
    PushI WL£y, WLn(Sq, R, SepChr)
Next
End Function
Private Function WLn$(Sq(), R&, SepChr$):      WLn = Join(DrStrSq(Sq, R), SepChr): End Function
Private Function WSqAli(Sq()) As Variant(): WSqAli = WSqAliWdty(Sq, WWdtySq(Sq)):  End Function
Private Function WSqAliWdty(Sq(), W%()) As Variant()
Dim O(): O = Sq
Dim C%: For C = 1 To NDcSq(Sq)
    Dim Wdt%: Wdt = W(C - 1)
    Dim IR&: For IR = 1 To NDrSq(Sq)
        O(IR, C) = AliL(O(IR, C), Wdt)
    Next
Next
WSqAliWdty = O
End Function
Private Function WWdtySq(Sq()) As Integer()
Dim C%: For C = 1 To UBound(Sq, 2)
    PushI WWdtySq, WWdtSqc(Sq, C)
Next
End Function
Private Function WWdtSqc%(Sq(), C%)
Dim R&, O%: For R = 1 To UBound(Sq, 1)
    O = Max(O, Len(Sq(R, C)))
Next
WWdtSqc = O
End Function
