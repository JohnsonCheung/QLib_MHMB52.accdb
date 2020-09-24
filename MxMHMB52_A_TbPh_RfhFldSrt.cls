VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_TbPh_RfhFldSrt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Sub RfhFldSrt()
DoCmd.SetWarnings False
Dim D As Dictionary: Set D = WDiPhToSno
With RsQC("SELECT Srt,Ph FROM ProdHierarchy where WithOHHst")
    While Not .EOF
        .Edit
        !Srt = WSrtPh(D, (Nz(!PH, "")))
        .Update
        .MoveNext
    Wend
End With
End Sub
Private Function WDiPhToSno() As Dictionary
Dim O As New Dictionary
With RsQC("Select Ph,format(Nz(x.Sno,0),'00') As Sno from ProdHierarchy x")
    While Not .EOF
        O.Add !PH.Value, !Sno.Value
        .MoveNext
    Wend
End With
Set WDiPhToSno = O
End Function
Private Function WSrtPh$(DiPhToSno As Dictionary, PH$)
Dim D As Dictionary: Set D = DiPhToSno
Select Case Len(PH)
Case 2: WSrtPh = D(PH)
Case 4: WSrtPh = D(Left(PH, 2)) & D(PH)
Case 7: WSrtPh = D(Left(PH, 2)) & D(Left(PH, 4)) & D(PH)
Case 10: WSrtPh = D(Left(PH, 2)) & D(Left(PH, 4)) & D(Left(PH, 7)) & D(PH)
End Select
End Function
