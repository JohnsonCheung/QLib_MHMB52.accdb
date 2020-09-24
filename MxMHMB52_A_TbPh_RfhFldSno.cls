VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_TbPh_RfhFldSno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Const CMod$ = "MxMHMB52_Ph_Rfh_TbProdHierarchy_FldSno."

Sub RfhFldSno()
MH.TbPh.RfhFldWithOHXxx
WUpdL1
WUpdL2ForL1
WUpdL3ForL2
WUpdL4ForL3
RunqC "Update ProdHierarchy set Sno=Null,Srt=Null where Not WithOHHst"
End Sub
Private Sub WUpdL1()
Dim Sno%
With RsQC("Select Sno from ProdHierarchy where WithOHHst and Lvl=1 order by Sno")
    While Not .EOF
        Sno = Sno + 1
        .Edit
        !Sno = Sno
        .Update
        .MoveNext
    Wend
End With
End Sub
Private Sub WUpdL2ForL1()
With RsQC("Select PH from ProdHierarchy where WithOHHst and Lvl=1")
    While Not .EOF
        WUpdL2 CStr(!PH)
        .MoveNext
    Wend
End With
End Sub
Private Sub WUpdL2(L1$)
Dim Sno%
With RsQC("Select Sno from ProdHierarchy where WithOHHst and Lvl=2 and Left(PH,2)='" & L1 & "' order by Sno")
    While Not .EOF
        Sno = Sno + 1
        .Edit
        !Sno = Sno
        .Update
        .MoveNext
    Wend
End With
End Sub

Private Sub WUpdL3ForL2()
With RsQC("Select PH from ProdHierarchy where WithOHHst and Lvl=2")
    While Not .EOF
        WUpdL3 CStr(!PH)
        .MoveNext
    Wend
End With
End Sub
Private Sub WUpdL3(L2$)
Dim Sno%
With RsQC("Select Sno from ProdHierarchy where WithOHHst and Lvl=3 and Left(PH,4)='" & L2 & "' order by Sno")
    While Not .EOF
        Sno = Sno + 1
        .Edit
        !Sno = Sno
        .Update
        .MoveNext
    Wend
End With
End Sub

Private Sub WUpdL4ForL3()
With RsQC("Select PH from ProdHierarchy where WithOHHst and Lvl=3")
    While Not .EOF
        WUpdL4 CStr(!PH)
        .MoveNext
    Wend
End With
End Sub
Private Sub WUpdL4(L3$)
Dim Sno%
With RsQC("Select Sno from ProdHierarchy where  WithOHHst and Lvl=4 and Left(PH,7)='" & L3 & "' order by Sno")
    While Not .EOF
        Sno = Sno + 1
        .Edit
        !Sno = Sno
        .Update
        .MoveNext
    Wend
End With
End Sub
