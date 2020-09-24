Attribute VB_Name = "MxVb_DDE"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_DDE."

Sub TstDDE()
Dim A As New Excel.Application
'Dim A As Excel.Application: Set A = GetObject(, "Excel.Application")
A.Visible = True
A.Workbooks.Add
If True Then
    Dim ChannelNumber&: ChannelNumber = Application.DDEInitiate( _
        Application:="Excel", _
        topic:="Book1")
    VBA.AppActivate "Book1"
    Application.DDEExecute ChannelNumber, "%{F11}"
    Application.DDETerminate ChannelNumber
    MsgBox "AA"
Else
    A.SendKeys "%{F11}"
End If
End Sub
