Attribute VB_Name = "MxDta_CnStr"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Cns."
'From:
'https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/sql/sql-server-express-user-instances
Public Const samp_Cns_SQLEXPR$ = "Data Source=.\\SQLExpress;Integrated Security=true;" & _
"User Instance=true;AttachDBFilename=|DataDirectory|\InstanceDB.mdf;" & _
"Initial Catalog=InstanceDB;"
'------------------------------------------
'From:
'https://social.msdn.microsoft.com/Forums/vstudio/en-US/61d45bef-eea7-4366-a8ad-e15a1fa3d544/vb6-to-connect-with-sqlexpress?forum=vbgeneral
Public Const samp_sqlExpr_notWrk_Cns3$ = _
"Provider=SQLNCLI.1;Integrated Security=SSPI;AttachDBFileName=C:\User\Users\northwnd.mdf;Data Source=.\sqlexpress"
Public Const Cns_ADO_SampSQL_EXPR_NOT_WRK$ = _
"Provider=OleDb;Integrated Security=SSPI;AttachDBFileName=C:\User\Users\northwnd.mdf;Data Source=.\sqlexpress"
'-----------------------------------------
'From https://social.msdn.microsoft.com/Forums/en-US/a73a838b-ec3f-419b-be65-8b1732fbf4d0/connect-to-a-remote-sql-server-db?forum=isvvba
Public Const samp_sqlExpr_notWrk_Cns1$ = "driver={SQL Server};" & _
      "server=LAPTOP-SH6AEQSO;uid=MyUserName;pwd=;database=pubs"
   
Public Const samp_sqlExpr_notWrk_Cns2$ = "driver={SQL Server};" & _
      "server=127.0.0.1;uid=MyUserName;pwd=;database=pubs"
   
Public Const samp_sqlExpr_notWrk_Cns$ = ".\SQLExpress;AttachDbFilename=c:\mydbfile.mdf;Database=dbname;" & _
"Trusted_Connection=Yes;"
'"Typical normal SQL Server connection string: Data Source=myServerAddress;
'"Initial Catalog=myDataBase;Integrated Security=SSPI;"

'From VisualStudio
Public Const SampSqlCns_NotWrk$ = _
    "Data Source=LAPTOP-SH6AEQSO\ProjectsV13;Initial Catalog=master;Integrated Security=True;Connect Timeout=30;" & _
    "Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"

Function CnsFbDao$(Fb): CnsFbDao = ";DATABASE=" & Fb & ";": End Function

Function CnsFxDao$(Fx)
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
'Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=C:\Users\sium\Desktop\TaxRate\sales text.xlsx;TABLE=Sheet1$
Dim O$
Select Case LCase(Ext(Fx))
Case ".xlsx":: O = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & Fx & ";"
Case ".xls": O = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & Fx & ";"
Case Else: Stop
End Select
CnsFxDao = O
End Function

Function CnsFcsvDao$(Fcsv)
Dim Fn$: Fn = Ffnn(Fcsv) & "#Csv"
CnsFcsvDao = FmtQQ("Text;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE=?;TABLE=?", Pth(Fcsv), Fn)
''Text;DSN=Delta_Tbl_08052203_20080522_033948 Link Specification;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE=C:\Tmp;TABLE=Delta_Tbl_08052203_20080522_033948#csv

End Function

Function Cnsy(D As Database) As String()
Dim T: For Each T In Tni(D)
    PushNB Cnsy, CnsT(D, T)
Next
End Function
