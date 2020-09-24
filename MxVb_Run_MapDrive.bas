Attribute VB_Name = "MxVb_Run_MapDrive"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Run_MapDrive."

Sub MapDrive(Drv$, Pth$)
RmvDrive Drv
Shell FmtQQ("Subst ? ""?""", Drv, Pth)
End Sub

Sub MapNDrive()
MapDrive "N:", "c:\users\user\desktop\Mhd"
End Sub

Sub RmvDrive(Drv$)
Shell "Subst /d " & Drv
End Sub

Sub RmvNDrive()
RmvDrive "N:"
End Sub
