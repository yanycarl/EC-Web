<%
Dim connString
Dim rs,rs2
int oid
set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs2=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
rs2.ActiveConnection = connString
oid=request.cookies("guango")("oid")
rs.Source ="update [order$] set situaction='�����Ѿ�����' where oid='"&oid&"'"
rs.open()
rs2.Source ="update [order$] set payment='��������' where oid='"&oid&"'"
rs2.open()
response.write("�����Ѿ���ʼ���ͣ����Ļ������������ѽ��ܡ�")
%>

<br><a href="personal.asp">���ظ�����ҳ</a></br>