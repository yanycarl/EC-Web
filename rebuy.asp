<center>
<br>
</br>
��Ʒ������
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

response.write("�û�ID��"+rs2("uid"))
response.write("<br>")
response.write("�û��ǳƣ�"+rs2("name"))
response.write("<br>")
response.write("��"+rs2("balance"))
response.write("<br>")


response.cookies("guango")("oid")=cstr(trim(Request("oid")))
response.write("<br><a href="&"pay.asp"&">��֧����֧��</a></br>")
response.write("<br><a href="&"payCharge.asp"&">�����֧��</a></br>")
response.write("<br><a href="&"payAfter.asp"&">��������</a></br>")


%>

<br>
<a href="default.html">������ҳ</a>
</br>
</center>