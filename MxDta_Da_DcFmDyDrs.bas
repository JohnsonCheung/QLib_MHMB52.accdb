Attribute VB_Name = "MxDta_Da_DcFmDyDrs"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_DcFmDyDrs."
Function DcBoolDrs(D As Drs, C) As Boolean(): DcBoolDrs = DcBoolDy(D.Dy, IxEle(D.Fny, C)): End Function
Function DcDrs(D As Drs, C) As Variant():         DcDrs = DcDy(D.Dy, IxEle(D.Fny, C)):     End Function
Function DcDblDrs(D As Drs, C) As Double():    DcDblDrs = DcDblDy(D.Dy, IxEle(D.Fny, C)):  End Function
Function DcIntDrs(D As Drs, C) As Integer():   DcIntDrs = DcIntDy(D.Dy, IxEle(D.Fny, C)):  End Function
Function DcLngDrs(D As Drs, C) As Long():      DcLngDrs = DcLngDy(D.Dy, IxEle(D.Fny, C)):  End Function
Function DcStrDrs(D As Drs, C) As String():    DcStrDrs = DcStrDy(D.Dy, IxEle(D.Fny, C)):  End Function

Function DcBoolDy(Dy(), C) As Boolean(): DcBoolDy = WWDcIntoDy(BoolyEmp, Dy, C): End Function
Function DcDy(Dy(), C) As Variant():         DcDy = WWDcIntoDy(AvEmp, Dy, C):    End Function
Function DcDblDy(Dy(), C) As Double():    DcDblDy = WWDcIntoDy(DblyEmp, Dy, C):  End Function
Function DcIntDy(Dy(), C) As Integer():   DcIntDy = WWDcIntoDy(DblyEmp, Dy, C):  End Function
Function DcLngDy(Dy(), C) As Long():      DcLngDy = WWDcIntoDy(LngyEmp, Dy, C):  End Function
Function DcStrDy(Dy(), C) As String():    DcStrDy = WWDcIntoDy(SyEmp, Dy, C):    End Function

Function DcFstDrs(D As Drs) As Variant():      DcFstDrs = DcFstDy(D.Dy):    End Function
Function DcStrFstDrs(D As Drs) As String(): DcStrFstDrs = DcStrDy(D.Dy, 0): End Function
Function DcFstDy(Dy()) As Variant():            DcFstDy = DcDy(Dy, 0):      End Function
Function DcStrSndDrs(D As Drs) As String(): DcStrSndDrs = DcStrDy(D.Dy, 1): End Function
Function DcStrSndDy(Dy()) As String():       DcStrSndDy = DcStrDy(Dy, 1):   End Function

Private Function WWDcIntoDy(Into, Dy(), C)
Dim O, U&
    U = UB(Dy)
    O = AyReDim(Into, U)
Dim Dr, J&: For Each Dr In Itr(Dy)
    If UB(Dr) >= C Then
        O(J) = Dr(C)
    End If
    J = J + 1
Next
WWDcIntoDy = O
End Function
