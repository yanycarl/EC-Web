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

response.write("订单支付失败，请充值。")
response.write("<br><a href="&"charge.asp"&">进入充值界面</a></br>")

else
rs.Source ="update [order$] set situaction='订单支付成功' where oid='"&request.cookies("guango")("oid")&"'"
rs.open()
rs4.Source ="update [consumer$] set balance=balance-"&rs2("prices")&" where uid='"&request.cookies("guango")("id")&"'"
rs4.open()
response.write("订单已经开始配送，您已支付成功。")

rs5.Source ="update [order$] set payment='余额支付' where oid='"&request.cookies("guango")("oid")&"'"
rs5.open()

end if
%>

<br><a href="personal.asp">返回个人主页</a></br>
</center>