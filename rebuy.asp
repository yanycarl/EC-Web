<center>
<br>
</br>
商品购买结果
<br>
</br>
<br>
</br>
<br>
</br>
<% 

Dim connString
Dim rs,rs2,rs3,rs4
string PRICE,name
int oid
int pid
Randomize Timer

set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs3=Server.CreateObject("ADODB.Recordset")
Set rs2=Server.CreateObject("ADODB.Recordset")
Set rs4=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
rs2.ActiveConnection = connString
rs3.ActiveConnection = connString
rs4.ActiveConnection = connString
        
        
response.write(request.cookies("guango")("name"))

rs2.Source ="select * from [consumer$] where uid='"&request.cookies("guango")("id")&"'"
rs2.Open()

response.write("用户ID："+rs2("uid"))
response.write("<br>")
response.write("用户昵称："+rs2("name"))
response.write("<br>")
response.write("余额："+rs2("balance"))
response.write("<br>")


response.cookies("guango")("oid")=cstr(trim(Request("oid")))
response.write("<br><a href="&"pay.asp"&">用支付宝支付</a></br>")
response.write("<br><a href="&"payCharge.asp"&">用余额支付</a></br>")
response.write("<br><a href="&"payAfter.asp"&">货到付款</a></br>")


%>

<br>
<a href="default.html">返回首页</a>
</br>
</center>