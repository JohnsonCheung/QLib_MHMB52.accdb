Attribute VB_Name = "MxDta_Da_Drs_Op_DrsJn"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Drs_Op_DrsJn."

Function DrsJnLeft(A As Drs, B As Drs, TmlJn$, CCExlB$) As Drs
DrsJnLeft = X_DrsJn(A, B, TmlJn, CCExlB, IsInrJn:=False)
End Function
Function DrsJnInr(A As Drs, B As Drs, TmlJn$, CCExlB$) As Drs
DrsJnInr = X_DrsJn(A, B, TmlJn, CCExlB, IsInrJn:=True)
End Function

Private Function X_DrsJn(A As Drs, B As Drs, TmlJn$, CCExlB$, IsInrJn As Boolean) As Drs
Dim FnyA$(), DyA()
Dim FnyB$(), DyB()
    FnyA = A.Fny: DyA = A.Dy
    FnyB = B.Fny: DyB = B.Dy

Dim DyKeyA(), DyKeyB()
    Dim FnyJnA$(), FnyJnB$()
    Dim CiyJnA%(), CiyJnB%()
    AsgTJn TmlJn, FnyJnA, FnyJnB
    CiyJnA = InyFnySub(FnyA, FnyJnA)
    CiyJnB = InyFnySub(FnyB, FnyJnB)
    DyKeyA = DySel(DyA, CiyJnA)
    DyKeyB = DySel(DyB, CiyJnB)
Dim UCol%: Stop '
Dim RxyA&(), RxyB&()
    WAsgRxyAB DyKeyA, DyKeyB, IsInrJn, RxyA, RxyB
Dim FnyO$()
Dim DyO()
    Dim CiySelB%(): Stop
    Dim DySelB()
        DySelB = DySel(DyB, CiySelB)
    DyO = WDyJn(DyA, DySelB, RxyA, RxyB, UCol)
    FnyO = SyAdd(FnyA, FnyB)
X_DrsJn = Drs(FnyO, DyO)
End Function
Private Sub WAsgRxyAB(DyKeyA(), DyKeyB(), IsInrJn As Boolean, ORxyA&(), ORxyB&())
Dim DrKeyA, RxyB&(), UbA&, UbB&
UbB = UB(DyKeyB)
UbA = UB(DyKeyA)
Dim RixA&, RixB&: For RixA = 0 To UB(DyKeyA)
    DrKeyA = DyKeyA(RixA)
    RxyB = WRxyHitB(DrKeyA, DyKeyB, UbB)
    Stop 'UHitB = UB(RxyB)
    Dim IHitB&: 'For IHitB = 0 To U
        PushI ORxyA, RixA
        'PushI ORxyB, RxyHitB(IHitB)
    'Next
    If Not IsInrJn Then
        'If UHitB = -1 Then
            PushI ORxyA, RixA
            PushI ORxyB, -1
        'End If
    End If
Next
End Sub
Private Function WRxyHitB(DrKeyA, DyKeyB(), UbB&) As Long()
Dim J&: For J = 0 To UbB
Stop '
    'If IsEqAy(DrKey, DyKeyB(J)) Then
        PushI WRxyHitB, J
    'End If
Next
End Function
Private Function WDyJn(DyA(), DySelB(), RxyA&(), RxyB&(), UCol%)
Dim J&: For J = 0 To UB(RxyA)
    Dim RixB&: RixB = RxyB(J)
    Dim RixA&: RixA = RxyA(J)
    Dim DrA: DrA = DyA(RixA)
    Dim DrI()
        If RxyB(J) = -1 Then
            DrI = DrA
            ReDim Preserve DrI(UCol)
        Else
            DrI = AyAdd(DrA, DySelB(RixB)) 'Assume all Dr in @DyA is same size.  So it is required ensured in the !Drs
        End If
    PushI WDyJn, DrI
Next
End Function
