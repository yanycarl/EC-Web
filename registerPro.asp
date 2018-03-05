<html>
<head>
<title>е§дкзЂВс</title>
</head>
<% 
Dim connString
Dim rs

dim name,price,pic,country
name=cstr(trim(Request("name")))
price=cstr(trim(Request("price")))
pic=cstr(trim(Request("pic")))
country=cstr(trim(Request("country")))


set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString

rs.Source ="insert into [newPro$](name,price,pic,country) values('"&name&"','"&price&"','"&pic&"','"&country&"')"
rs.Open()

response.redirect("AdminInfo.asp")

%>
</html>