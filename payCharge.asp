<center>
<%
Dim connString
Dim rs,rs2,rs3,rs4
dim a,b
set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs2=Server.CreateObject("ADODB.Recordset")
Set rs3=Server.CreateObject("ADODB.Recordset")
Set rs4=Server.CreateObject("ADODB.Recordset")
Set rs5=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
rs2.ActiveConnection = connString
rs3.ActiveConnection = connString
rs4.ActiveConnection = connString

rs5.ActiveConnection = connString

rs2.Source ="select * from [order$] where oid='"&request.cookies("guango")("oid")&"'"
rs2.open()

rs3.Source ="select * from [consumer$] where uid='"&request.cookies("guango")("id")&"'"
rs3.open()

if cint(rs2("prices")) > cint(rs3("balance")) then

response.write("����֧��ʧ�ܣ����ֵ��")
response.write("<br><a href="&"charge.asp"&">�����ֵ����</a></br>")

else
rs.Source ="update [order$] set situaction='����֧���ɹ�' where oid='"&request.cookies("guango")("oid")&"'"
rs.open()
rs4.Source ="update [consumer$] set balance=balance-"&rs2("prices")&" where uid='"&request.cookies("guango")("id")&"'"
rs4.open()
response.write("�����Ѿ���ʼ���ͣ�����֧���ɹ���")

rs5.Source ="update [order$] set payment='���֧��' where oid='"&request.cookies("guango")("oid")&"'"
rs5.open()

end if
%>

<br><a href="personal.asp">���ظ�����ҳ</a></br>
</center>