Attribute VB_Name = "MxXls_TreeInstall"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_TreeInstall."
Function CdlTreeWs$()
Erase XX
X "Option Explicit"
X "Sub Worksheet_Change(ByVal Target As Range)"
X "MTreeWs.Change Target"
X "End Sub"

X "Sub Worksheet_SelectionChange(ByVal Target As Range)"
X "MTreeWs.SelectionChange Target"
X "End Sub"
CdlTreeWs = JnCrLf(XX)
Erase XX
End Function

Sub InstallTreeWs()
Dim Ws, Wb
For Each Ws In Itr(WsyTree)
    WInstall CvWs(Ws)
Next

Stop 'For Each Wb In Itr(WbyTree)
    'InstallTreeWb CvWb(Wb)
'Next
End Sub
Private Sub WInstall(S As Worksheet)
Const CSub$ = CMod & "WInstall"
Dim Md As CodeModule
Set Md = MdWs(S)
Stop
If Md.CountOfLines = 0 Then
    Md.AddFromString CdlTreeWs
    Inf CSub, "TreeWs in Wb is installed with code", "Wb", WbnWs(S)
Else
    Inf CSub, "TreeWs in Wb already has code", "Wb", WbnWs(S)
End If
End Sub

Function IsWbTree(B As Workbook) As Boolean
Dim Ws As Worksheet
For Each Ws In B.Sheets
    If Ws.Name = "TreeWs" Then IsWbTree = True: Exit Function
Next
End Function
Function WbyTree() As Workbook()
Dim Wb As Workbook, Ws As Worksheet
For Each Wb In Xls.Workbooks
    If IsWbTree(Wb) Then PushObj WbyTree, Wb
Next

End Function
Function WsyTree() As Worksheet()
Dim Wb As Workbook, Ws As Worksheet
For Each Wb In Xls.Workbooks
    For Each Ws In Wb.Sheets
        If Ws.Name = "TreeWs" Then PushObj WsyTree, Ws
    Next
Next
End Function
