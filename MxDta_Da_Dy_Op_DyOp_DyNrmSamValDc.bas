Attribute VB_Name = "MxDta_Da_Dy_Op_DyOp_DyNrmSamValDc"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dy_Op_DyOp_DyNrmSamValDc."
Private Sub B_DyNrmEqvDc()
GoSub T1
GoSub T2
GoSub T3
Exit Sub
Dim Dy(), NDc%, E, Act(), Ept()
T1:
    Dy = Array( _
        Array(1, 2, 3, 4, 5), _
        Array(1, 2, 3, 4, 5), _
        Array(1, 2, 3, 4, 5))
    Ept = Array( _
        Array(1, 2, 3, 4, 5), _
        Array(E, E, 3, 4, 5), _
        Array(E, E, 3, 4, 5))
    NDc = 2
    GoTo Tst
    Return
T2:
    Dy = Array( _
        Array(1, 2, 3, 4, 5), _
        Array(1, 2, 3, 4, 5), _
        Array(1, 2, 3, 4, 5), _
        Array(1, 3, 3, 4, 5))
    Ept = Array( _
        Array(1, 2, 3, 4, 5), _
        Array(E, E, 3, 4, 5), _
        Array(E, E, 3, 4, 5), _
        Array(E, 3, 3, 4, 5))
T3:
    Dy = Array( _
        Array(1, 2, 3, 4, 5), _
        Array(1, 2, 3, 4, 5), _
        Array(1, 2, 3, 4, 5), _
        Array(2, 2, 3, 4, 5))
    Ept = Array( _
        Array(1, 2, 3, 4, 5), _
        Array(E, E, 3, 4, 5), _
        Array(E, E, 3, 4, 5), _
        Array(2, 2, 3, 4, 5))
    GoTo Tst
Tst:
    Act = DyNrmEqvDc(Dy, NDc)
    C
    Return
End Sub
Function DyNrmEqvDc1(Dy()) As Variant(): DyNrmEqvDc1 = DyNrmEqvDc(Dy, 1): End Function
Function DyNrmEqvDc2(Dy()) As Variant(): DyNrmEqvDc2 = DyNrmEqvDc(Dy, 2): End Function
Function DyNrmEqvDc3(Dy()) As Variant(): DyNrmEqvDc3 = DyNrmEqvDc(Dy, 3): End Function
Function DyNrmEqvDc(Dy(), NDc%) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim DrLas(): DrLas = AwFstN(Dy(0), NDc)
Dim ODy()
    ODy = Dy
    Dim IDr&: For IDr = 1 To UB(Dy)
        Dim DrCur(): DrCur = AwFstN(Dy(IDr), NDc)
        Dim B() As Boolean: B = WBoolyDcShdHid(DrLas, DrCur)
        Dim IDc%: For IDc = 0 To NDc - 1
            If B(IDc) Then ODy(IDr)(IDc) = Empty
        Next
        DrLas = DrCur
    Next
DyNrmEqvDc = ODy
End Function
Private Function WBoolyDcShdHid(DrLas(), DrCur()) As Boolean()
Dim UDc%: UDc = UB(DrLas)
Dim O() As Boolean: ReDim O(UDc)
Dim IDc%: For IDc = 0 To UDc
    If Not IsEqAyFstN(DrLas, DrCur, IDc) Then GoTo E
    O(IDc) = True
Next
E: WBoolyDcShdHid = O
End Function
