Attribute VB_Name = "MxDao_Def_TdPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Def_TdPrp."

Sub SetPvDesDi(D As Database, TblDes As Dictionary)
Dim T: For Each T In TblDes.Keys
    SetPvTDes D, T, TblDes(T)
Next
End Sub

Sub SetFDes(D As Database, DiFDes As Dictionary)
Dim TF: For Each TF In DiFDes.Keys
    Dim T$, F$
    With BrkDot(TF)
        T = .S1
        F = .S2
    End With
    SetPvFDes D, T, F, DiFDes(TF)
Next
End Sub

Function DiTDes(D As Database) As Dictionary
Dim T, O As New Dictionary
For Each T In Tni(D)
    PushKvNBDrp O, T, PvTDes(D, T)
Next
Set DiTDes = O
End Function

Sub SetTbPrp(D As Database, T, P$, V)

End Sub

Function DaoPvP(P As Dao.Properties, Pn$)
If HasPrp(P, Pn) Then DaoPvP = P(Pn).Value
End Function

Function DaoPrps(DaoPrpsObj) As Dao.Properties
Set DaoPrps = DaoPrpsObj.Properties
End Function

Function DaoPv(DaoPrpObj, P$)
DaoPv = DaoPvP(DaoPrpObj.Properties, P)
End Function
