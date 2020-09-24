Attribute VB_Name = "MxIde_Src_Cpmd_CpmdOp"
Option Compare Text
Option Explicit
Function CpmdlyCpmdy(A() As Cpmd) As Cpmdl()
Dim J&: For J = 0 To UbCpmd(A)
    PushCpmdl CpmdlyCpmdy, CpmdlCpmd(A(J))
Next
End Function
Function CpmdlCpmd(A As Cpmd) As Cpmdl
With CpmdlCpmd
    .Mdn = A.Mdn
    .Befl = JnCrLf(A.Bef)
    .Aftl = JnCrLf(A.Aft)
End With
End Function
Function CpmdlyWhDif(A() As Cpmdl) As Cpmdl()
Dim J&: For J = 0 To UbCpmdl(A)
    With A(J)
        If .Aftl <> .Befl Then
            PushCpmdl CpmdlyWhDif, A(J)
        End If
    End With
Next
End Function

Function CpmdyMsrcy(Msrcy() As Nly) As Cpmd()
Dim J&: For J = 0 To UbNly(Msrcy)
    With Msrcy(J)
        Dim Bef$(), Aft$()
        Bef = LyEndTrim(SrcMdn(.Nm))
        Aft = LyEndTrim(.Ly)
        If Not IsEqAy(Bef, Aft) Then
            PushCpmd CpmdyMsrcy, Cpmd(.Nm, Bef, Aft)
        End If
    End With
Next
End Function
