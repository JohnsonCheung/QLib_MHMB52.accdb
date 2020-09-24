Attribute VB_Name = "MxDta_Da_Samp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Samp."
Function sampDr1() As Variant(): sampDr1 = Array(1, 2, 3): End Function

Function sampDr2() As Variant():  sampDr2 = Array(2, 3, 4):                       End Function
Function sampDr3() As Variant():  sampDr3 = Array(3, 4, 5):                       End Function
Function sampDr4() As Variant():  sampDr4 = Array(43, 44, 45):                    End Function
Function sampDr5() As Variant():  sampDr5 = Array(53, 54, 55):                    End Function
Function sampDr6() As Variant():  sampDr6 = Array(63, 64, 65):                    End Function
Function sampDrs2() As Drs:      sampDrs2 = DrsFf("A B C", sampDy2):              End Function
Function sampDrs1() As Drs:      sampDrs1 = DrsFf("A B C", sampDy1):              End Function
Function sampDrs() As Drs:        sampDrs = DrsFf("A B C D E G H I J K", sampDy): End Function
Function sampDy1() As Variant():  sampDy1 = Array(sampDr1, sampDr2, sampDr3):     End Function
Function sampDy2() As Variant(): sampDy2 = Array(sampDr3, sampDr4, sampDr5)
End Function

Function sampDy3() As Variant()
PushI sampDy3, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(100, "A") & vbCrLf & String(100, "X"))
PushI sampDy3, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(100, "A") & vbCrLf & String(100, "X"))
End Function

Function sampDy() As Variant()
PushI sampDy, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "A"))
PushI sampDy, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "B"))
PushI sampDy, Array("C", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "C"))
PushI sampDy, Array("D", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "D"))
PushI sampDy, Array("E", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "E"))
PushI sampDy, Array("F", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "F"))
PushI sampDy, Array("G", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "G"))
End Function

Function sampDs() As Ds
PushDt sampDs.Dty, sampDt1
PushDt sampDs.Dty, sampDt2
sampDs.Dsn = "Ds"
End Function

Function sampDt1() As Dt: sampDt1 = DtFf("SampDt1", "A B C", sampDy1): End Function
Function sampDt2() As Dt: sampDt2 = DtFf("SampDt2", "A B C", sampDy2): End Function

Function sampDrAToJ() As Variant()
Const NC% = 10
Dim J%
For J = 0 To NC - 1
    PushI sampDrAToJ, Chr(Asc("A") + J)
Next
End Function
