Attribute VB_Name = "MxXls RgPr_HypLnk"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_HypLnk."

Private Sub B_HypLnkRgPr()
Dim A As Range, B As Range, S As Worksheet
GoSub ZZ
Exit Sub
ZZ:
    Set S = WsNw
    WsAdd WbWs(S)
    Set A = S.Range("A1")
    Set B = WbWs(S).Sheets(2).Range("A1")
    HypLnkRgPr A, B
    Maxv S.Application
    Return
Crt:
    Return
End Sub
Sub HypLnkRgPr(A As Range, B As Range, Optional ByVal DspA$, Optional ByVal DspB$) ' Set HLnk to a pair of cell
Const CSub$ = CMod & "HypLnkRgPr"
ChkIsCell A, CSub
ChkIsCell B, CSub
If DspA = "" Then If IsEmpty(A.Value) Then DspA = WsRg(B).Name Else DspA = A.Value
If DspB = "" Then If IsEmpty(B.Value) Then DspB = WsRg(A).Name Else DspB = B.Value
DltHypLnk A
DltHypLnk B
WsRg(A).Hyperlinks.Add Anchor:=A, Address:="", SubAddress:=AdrRg(B), TextToDisplay:=DspA
WsRg(B).Hyperlinks.Add Anchor:=B, Address:="", SubAddress:=AdrRg(A), TextToDisplay:=DspB
End Sub
Sub DltHypLnk(R As Range)
If Not IsCell(R) Then Exit Sub
On Error Resume Next
R.Hyperlinks.Delete
End Sub
