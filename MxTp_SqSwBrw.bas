Attribute VB_Name = "MxTp_SqSwBrw"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_SqSwBrw."
Private Function FmtSqSw(A As SqSw) As String()
ClrXX
X SyAddAp(Box("Pm"), FmtDi(A.fldSw.Pm))
X SyAddAp(Box("stmtSw"), FmtDi(A.stmtSw))
X SyAddAp(Box("fldSw"), FmtDi(A.fldSw))
FmtSqSw = XX
End Function
Private Sub BrwSwly(L() As Swl): BrwAy FmtSwly(L): End Sub
Function FmtSwly(L() As Swl) As String()
Dim J%: For J = 0 To SwlUB(L)
    PushI FmtSwly, FmtSwl(L(J))
Next
End Function
Function FmtSwl$(L As Swl)
With L
Dim X$
Select Case True
Case IsOrAndStr(.Op): X = JnSpc(.Tml)
Case IsEqNeStr(.Op): X = .Tml(0) & " " & .Tml(1)
End Select
FmtSwl = JnSpcAp(.Swn, CvBoolOp(.Op), X)
End With
End Function
