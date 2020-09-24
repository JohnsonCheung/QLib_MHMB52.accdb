Attribute VB_Name = "MxIde_Mthn_Cml_Mi2AdjelFunnn"
Option Compare Database
Option Explicit
Private Sub VcMi2AdjelFunnn__Tst():     VcMi2AdjelFunnn:                        End Sub
Sub VcMi2AdjelFunnn(Optional W% = 130): Vc Mi2yAdjelFunnnPC(W), "Adjel Funnn ": End Sub
Private Function Mi2yAdjelFunnnPC(Optional W% = 130) As String(): Mi2yAdjelFunnnPC = FmtParcc(Mi2yAdjelFunnPC, H12:="Adjel Funn", Wdt:=W): End Function
Private Sub Mi2yAdjelFunnPC__Tst(): Vc FmtT1ry(Mi2yAdjelFunnPC), "Adjel Fun": End Sub
Function Mi2yAdjelFunnPC() As String()
Dim Funn: For Each Funn In FunnyPubPC
    Dim Adjel$: Adjel = MiAdjelzFunn(Funn)
    If Adjel <> "" Then
        PushNB Mi2yAdjelFunnPC, Adjel & " " & Funn
    End If
Next
End Function
