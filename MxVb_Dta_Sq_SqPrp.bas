Attribute VB_Name = "MxVb_Dta_Sq_SqPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Sq_Prp."
Function DcStrSq(Sq, Optional C = 1) As String():  DcStrSq = WIntoSqc(SyEmp, Sq, C):   End Function ' Sq uses variant, because it allows String or Variant Sq
Function DcIntSq(Sq, Optional C = 1) As Integer(): DcIntSq = WIntoSqc(IntyEmp, Sq, C): End Function ' Sq uses variant, because it allows String or Variant Sq
Function DcSq(Sq, Optional C = 1) As Variant():       DcSq = WIntoSqc(AvEmp, Sq, C):   End Function
Function DrSq(Sq, Optional R = 1) As Variant():       DrSq = WIntoSqr(AvEmp, Sq, R):   End Function
Function DrStrSq(Sq, Optional R = 1) As String():  DrStrSq = WIntoSqc(SyEmp, Sq, R):   End Function
Function SySqr(Sq, Optional R = 1) As String():      SySqr = WIntoSqr(SyEmp, Sq, R):   End Function
Function SySqc(Sq, Optional C = 1) As String():      SySqc = WIntoSqc(SyEmp, Sq, C):   End Function

Function NDrSq&(Sq)
On Error Resume Next
NDrSq = UBound(Sq, 1)
End Function
Function NDcSq&(Sq)
On Error Resume Next
NDcSq = UBound(Sq, 2)
End Function

Function DrSqCnoy(Sq(), R, Cnoy%()) As Variant(): DrSqCnoy = WDrInto(AvEmp, Sq, R, Cnoy): End Function
Private Function WDrInto(Into, Sq(), R, Cnoy)
Dim UCol%:    UCol = UBound(Cnoy)
Dim O: O = Into: ReDim O(UCol)
Dim C%: For C = 1 To UCol + 1
    O(C - 1) = Sq(R, C)
Next
WDrInto = O
End Function

Private Function WIntoSqc(Into, Sq, C)
Dim NR&: NR = UBound(Sq, 1)
Dim O:    O = AyReDim(Into, NR - 1)
Dim R&: For R = 1 To NR
    O(R - 1) = Sq(R, C)
Next
WIntoSqc = O
End Function
Private Function WIntoSqr(Into, Sq, R)
Dim NCol&:    NCol = UBound(Sq, 2)
Dim O: O = Into: ReDim O(NCol - 1)
Dim C%: For C = 1 To NCol
    O(C - 1) = Sq(R, C)
Next
WIntoSqr = O
End Function
