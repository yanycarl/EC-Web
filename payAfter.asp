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
rs.Source ="update [order$] set situaction='货物已经发送' where oid='"&oid&"'"
rs.open()
rs2.Source ="update [order$] set payment='货到付款' where oid='"&oid&"'"
rs2.open()
response.write("订单已经开始配送，您的货到付款请求已接受。")
%>

<br><a href="personal.asp">返回个人主页</a></br>