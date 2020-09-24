Attribute VB_Name = "MxDao_TblNRec"
Option Compare Text
Option Explicit

Const CMod$ = "MxDao_TblNRec."
Enum eSrtTbNRec: eSrtTbNRecByN: eSrtTbNRecByTbn: End Enum: Public Const EnmmSrtTbNRec$ = "eSrtTbNRec? ByN ByTbn"
Sub DmpTbNRecC(Optional S As eSrtTbNRec):               DmpTbNRec CDb, S:           End Sub
Sub BrwTbNRecC(Optional S As eSrtTbNRec):               BrwTbNRec CDb, S:           End Sub
Sub VcTbNRecC(Optional S As eSrtTbNRec):                VcTbNRec CDb, S:            End Sub
Sub DmpTbNRec(D As Database, Optional S As eSrtTbNRec): Dmp FmtTbNRec(D, S):        End Sub
Sub VcTbNRec(D As Database, Optional S As eSrtTbNRec):  Vc FmtTbNRec(D, S):         End Sub
Sub BrwTbNRec(D As Database, Optional S As eSrtTbNRec): Stop 'BrwT FmtTbNRec(D, S): End Sub

End Sub
Function FmtTbNRec(D As Database, S As eSrtTbNRec) As String()
Const CSub$ = CMod & "FmtNRec"
Dim T$(): T = Tny(D)
ClrBfr
Dim I: For Each I In Itr(T)
    BfrV NRecT(D, I) & " " & I
Next
Stop 'PushI FmtTbNRec, SymNap("NTbl Fb", D.Name, Si(T))
Stop 'Dim S$: S = WHypKeyCii(S)
Stop 'PushIAy FmtTbNRec, FmtT1ry(LyBfr, "0", S, "0")
End Function
Private Function WHypKeyCii$(S As eSrtTbNRec)

End Function
