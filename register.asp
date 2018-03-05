<html>
<head>
<title>正在注册</title>
</head>
<% 
Dim connString
Dim rs

dim username,password,name,addr,sex,region,tel,email
username=cstr(trim(Request("username")))
password=cstr(trim(Request("password")))
name=cstr(trim(Request("name")))
addr=cstr(trim(Request("addr")))
sex=cstr(trim(Request("sex")))
region=cstr(trim(Request("region")))
tel=cstr(trim(Request("tel")))
email=cstr(trim(Request("email")))

set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString

if username <> "" or password <> "" then
rs.Source ="insert into [consumer$](uid,pwd,name,balance,sid,living,addr,sex,region,sign,tel,email) values('"&username&"','"&password&"','"&name&"','0','"&username&"','0','"&addr&"','"&sex&"','"&region&"','0','"&tel&"','"&email&"')"
rs.Open()

response.redirect("login.html")
else
response.write("注册失败,请重新注册。")
response.redirect("register.html")
end if
%>
</html>