Attribute VB_Name = "MxDta_Da_Wh_DwDe_DeDup"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Wh_DwDe_DeDup."

Function DwDup(D As Drs, CC$) As Drs: DwDup = DwRxy(D, RxyDupDrs(D, CC)): End Function

Function DeDup(D As Drs) As Drs
DeDup = DeDupFf(D, JnSpc(D.Fny))
End Function

Function DeDupFf(D As Drs, CCDup$) As Drs
Dim Rxy&(): Rxy = RxyDupDrs(D, CCDup)
DeDupFf = DeRxy(D, Rxy)
End Function

Private Function RxyDupDy(Dy()) As Long()
Dim DyDup(): DyDup = DyWhDup(Dy)
Dim Dr, Ix&, O&()
For Each Dr In Dy
    If HasDr(DyDup, Dr) Then PushI O, Ix
    Ix = Ix + 1
Next
If Si(O) < Si(DyDup) * 2 Then Stop
RxyDupDy = O
End Function

Private Function RxyDupDrs(D As Drs, CC$) As Long()
Dim Ciy%(): Ciy = InyDrsCc(D, CC)
If Si(Ciy) = 1 Then
    RxyDupDrs = IxyDup(DcDy(D.Dy, Ciy(0)))
    Exit Function
End If
Dim Dy(): Dy = DySel(D.Dy, Ciy)
RxyDupDrs = RxyDupDy(Dy)
End Function

Function DyWhDup(Dy()) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim Dr: For Each Dr In DyGpCnt(Dy)
    If Dr(0) > 1 Then
        PushI DyWhDup, AeFst(Dr)
    End If
Next
End Function

Function DyWhDupCiy(Dy(), Ciy&()) As Variant()
Dim Dup$(), Dr
Dup = AwDup(DyJn(Dy, Ciy))
For Each Dr In Itr(Dy)
    If HasEle(Dup, Jn(AwIxy(Dr, Ciy), vbFldSep)) Then Push DyWhDupCiy, Dr
Next
End Function

Function DyWhDupC(Dy(), C&) As Variant()
DyWhDupC = AwIxy(Dy, IxyDup(DcDy(Dy, C)))
End Function
