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

if request.cookies("guango")("id")="0" then

    response.redirect("login.html")
else

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


rs3.Source ="select * from [products$] where name='"&request.cookies("guango")("name")&"'"
rs3.open()
if rs3.eof or rs3.bof then
response.write("���޴���Ʒ")
else
response.write("��Ʒ����"+rs3("name"))
response.write("<br>")
response.write("���ۣ�"+rs3("price"))
response.write("<br>")
response.write("���أ�"+rs3("country"))
response.write("<br>")


price=rs3("price")
name=rs3("name")
oid=Int(899999 * Rnd + 100000)
pid=Int(899999 * Rnd + 100000)

rs.Source ="insert into [order$] values('"&rs2("uid")&"','"&oid&"','0','"&name&"','"&price&"','1','��ַδ��д','"&price&"','�����ɹ���δ֧��','"&pid&"','δѡ��֧����ʽ');"
rs.Open()


rs4.Source ="select * from [order$] where oid='"&oid&"'"
rs4.open()


response.write("�����ţ�"+rs4("oid"))
response.write("<br>")
response.write("״̬��"+rs4("situaction"))
response.write("<br>")
response.cookies("guango")("oid")=rs4("oid")
response.write("<br><a href="&"pay.asp"&">��֧����֧��</a></br>")
response.write("<br><a href="&"payCharge.asp"&">�����֧��</a></br>")
response.write("<br><a href="&"payAfter.asp"&">��������</a></br>")

end if
end if

%>

<br>
<a href="default.html">������ҳ</a>
</br>
</center>