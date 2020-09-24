Attribute VB_Name = "MxVb_Dta_Asc"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Asc."

Function AscCnty(S) As Long() ' It an long-ay of 256 elements of a sum.  Each element respresents an ascii.  Its value of is the count of that asc of @S.
' And Len(S) should always equal to Sum of the @@return.
Dim O&(255), A As Byte
Dim J&
For J = 1 To Len(S)
    A = Asc(Mid(S, J, 1))
    O(A) = O(A) + 1
Next
AscCnty = O
End Function
