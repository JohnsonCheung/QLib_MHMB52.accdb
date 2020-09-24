Attribute VB_Name = "MxVb_FmtColQuo"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_FmtColQuo."
Type Qmk: Q1 As String: Sep As String: Q2 As String: End Type 'Deriving(Ctor) #Quo-Mrk-For-a-Line#
Function Qmk(Q1, Sep, Q2) As Qmk
With Qmk
    .Q1 = Q1
    .Sep = Sep
    .Q2 = Q2
End With
End Function

Function QmkSepln(F As eTblFmt) As Qmk
Const CSub$ = CMod & "QmkSepln"
If F = eTblFmtSS Then
    QmkSepln = QmkSeplnSS
ElseIf F = eTblFmtTb Then
    QmkSepln = QmkSeplnTB
Else
    ThwEnm CSub, F, EnmqssTblFmt
End If
End Function
Function QmkDta(F As eTblFmt) As Qmk
Const CSub$ = CMod & "QmkDta"
If F = eTblFmtSS Then
    QmkDta = QmkDtaSS
ElseIf F = eTblFmtTb Then
    QmkDta = QmkDtaTB
Else
    ThwEnm CSub, F, EnmqssTblFmt
End If
End Function
Function QmkSeplnTB() As Qmk: With QmkSeplnTB: .Sep = "-|-": .Q1 = "|-": .Q2 = "-|": End With: End Function
Function QmkSeplnSS() As Qmk: QmkSeplnSS = QmkDtaSS: End Function
Function QmkDtaTB() As Qmk: With QmkDtaTB: .Sep = " | ": .Q1 = "| ": .Q2 = " |": End With: End Function
Function QmkDtaSS() As Qmk: With QmkDtaSS: .Sep = " ":                           End With: End Function
