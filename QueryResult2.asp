<center>
<br><font size="+3">��ѯ���:</font></br>
<br></br>
<%


who=cstr(trim(Request("op")))
	
if request.cookies("guango").haskeys then '�����Ƿ���Cookie

Dim connString
Dim rs,rs2 '��������
Dim oidString
Dim who

set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs2=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
rs2.ActiveConnection = connString '�������ݿⲢ��ʼ����¼��


rs.Source =""&who&""
rs.Open()


do while not rs.bof or not rs.eof



    if rs.eof or rs.bof then
 
    response.write("<br>")
    response.write("û����������")

   response.write("<br>")
    response.write("<br>")    
response.write("<br>")
    response.write("<br>")    
response.write("<br>")
    response.write("<br>")    
response.write("<br>")
    response.write("<br>")    
response.write("<br>")
    response.write("<br>")    
response.write("<br>")
    response.write("<br>")    
response.write("<br>")
    response.write("<br>")    
response.write("<br>")
    response.write("<br>")
    else 
    response.write(" �û�ID��")
    response.write(rs("uid"))
    response.write("    �����ţ�")
    response.write(rs("oid"))
    response.write("   ��Ʒ����")
    response.write(rs("name"))
    response.write("   ���ۣ�")
    response.write(rs("price"))
    response.write("   ������")
    response.write(rs("amount"))
    response.write("   �ܼ۸�")
    response.write(rs("prices"))
    response.write("   ���ʽ��")
    response.write(rs("payment"))
    response.write("   ��ݵ��ţ�")
    response.write(rs("expressinfo"))
    response.write("   ״̬��")
    response.write(rs("situaction"))
    if not rs("situaction")="����֧���ɹ�" then
    end if

end if
'"&oidString&"
    response.write("<br>")
    response.write("<br>")

rs.movenext

loop

 
    
else 
    response.redirect(login.html)
end if
%>
</center>