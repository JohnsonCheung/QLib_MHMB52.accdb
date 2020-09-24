Attribute VB_Name = "MxAcs_Acs_MMaxAcs"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_Acs_MiMaxAcs."

Sub MinvAcs(A As Access.Application): VisAcs A: MiniAcs A:                 End Sub
Sub MaxvAcs(A As Access.Application): VisAcs A: MiniAcs A:                 End Sub
Sub MiniAcs(A As Access.Application): A.DoCmd.RunCommand acCmdAppMinimize: End Sub
Sub MaxiAcs(A As Access.Application): A.DoCmd.RunCommand acCmdAppMaximize: End Sub
Sub VisAcs(A As Access.Application):  A.Visible = True:                    End Sub
