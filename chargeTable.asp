<center>
<br>
</br>
屎壓割峙。。。
<br>
</br>
<br>
</br>
<br>
</br>
<% 
dim b

b=cstr(trim(Request("balance")))

Dim connString
Dim rs

set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString

	rs.Source ="update [consumer$] set balance=balance+"&b&" where uid='"&request.cookies("guango")("id")&"'"
	rs.Open()
	

response.redirect("personal.asp")
%>
</center>