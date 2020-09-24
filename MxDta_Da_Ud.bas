Attribute VB_Name = "MxDta_Da_Ud"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Ud."

Type Drs: Fny() As String: Dy() As Variant: End Type 'Deriving(Ctor)
Type Dt: Dtn As String: Fny() As String: Dy() As Variant: End Type 'Deriving(Ctor)
Type Rec: Fny() As String: Dr() As Variant: End Type 'Deriving(Ctor)
Type Ds: Dsn As String: Dty() As Dt: End Type ' Deriving(Ctor)
Function Drs(Fny$(), Dy() As Variant) As Drs
With Drs
    .Fny = Fny
    .Dy = X_DyEnsCol(Dy, UB(Fny))
End With
End Function
Function Dt(Dtn, Fny$(), Dy() As Variant) As Dt
With Dt
    .Dtn = Dtn
    .Fny = Fny
    .Dy = X_DyEnsCol(Dy, UB(Fny))
End With
End Function
Function Rec(Fny$(), Dr() As Variant) As Rec
With Rec
    .Fny = Fny
    .Dr = Dr
End With
End Function
Function Ds(Dsn, Dty() As Dt) As Ds
With Ds
    .Dsn = Dsn
    .Dty = Dty
End With
End Function

Private Function X_DyEnsCol(Dy(), UCol%) As Variant()
Const CSub$ = CMod & "X_DyEnsCol"
Dim O(): O = Dy
Dim J&: For J = 0 To UB(Dy)
    Dim UI&: UI = UB(Dy(J))
    Select Case UI
    Case Is = UCol
    Case Is > UCol: Thw CSub, "The *J-th Dr of @Dy has UB > @UCol", "*J UB(Dy(J)) UCOl", J, UI, UCol
    Case Else: ReDim Preserve O(J) '<===
    End Select
Next
X_DyEnsCol = O
End Function
