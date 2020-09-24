Attribute VB_Name = "MxDta_Da_Wh_DyWh"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Wh_DyWh."

Function DyWhDr(Dy(), DrWh(), DyKey()) As Variant()
DyWhDr = AwIxy(Dy, WRxyDyWhDr(DyKey, DrWh))
End Function

Private Function WRxyDyWhDr(Dy(), WhDr()) As Long()
Dim Dr, I&: For Each Dr In Itr(Dy)
    If IsEqDr(Dr, WhDr) Then
        #If False Then
        Stop
        Debug.Print I
        Debug.Print JnSpc(Dr)
        Debug.Print JnSpc(WhDr)
        Debug.Print
        #End If
        PushI WRxyDyWhDr, I
    End If
    I = I + 1
Next
End Function

Private Sub B_WRxyDyWhDr()
GoSub T1
Exit Sub
Dim DrWh(), Dy()
T1:
    Dy = DrsSelFf(DrsTMdPC, "Mdn").Dy
    Stop
    DrWh = Array("QGit", Empty)
    GoTo Tst
T2:
    Dy = DrsSelFf(DrsTMdPC, "Mdn").Dy
    DrWh = Array("QAct", Empty)
    GoTo Tst
Tst:
    Act = WRxyDyWhDr(DrWh, Dy)
    Dmp Act
    Return
End Sub
