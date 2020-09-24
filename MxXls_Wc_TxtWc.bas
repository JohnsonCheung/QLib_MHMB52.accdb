Attribute VB_Name = "MxXls_Wc_TxtWc"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Wc_TxtWc."

Function WcTxt(C As WorkbookConnection) As TextConnection
On Error Resume Next
Set WcTxt = C.TextConnection
End Function

Function WcyTxtzWb(B As Workbook) As TextConnection()
Dim C As WorkbookConnection: For Each C In B.Connections
    If Not IsNothing(WcTxt(C)) Then
        PushObj WcyTxtzWb, C.TextConnection
    End If
Next
End Function

Private Sub B_NWcTxt()
Dim O As Workbook: Set O = samp_mhmb52rpt_Wb
Ass NWcTxt(O) = 0
O.Application.Quit
End Sub
Function NWcTxt%(B As Workbook)
Dim C As WorkbookConnection: For Each C In B.Connections
    If Not IsNothing(WcTxt(C)) Then NWcTxt = NWcTxt + 1
Next
End Function

Function CnsyWcTxtzWb(B As Workbook) As String()
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Stop 'Dim T As TextConnection: Set T = TxtWcCnsy_Wb(B)
'If IsNothing(T) Then Exit Function
'TxtWcCnsy_Wb = T.Connection
End Function
