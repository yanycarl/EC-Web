<center>
<br><font size="+3">查询结果:</font></br>
<br></br>
<%


who=cstr(trim(Request("op")))
	
if request.cookies("guango").haskeys then '测试是否有Cookie

Dim connString
Dim rs,rs2 '声明对象
Dim oidString
Dim who

set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs2=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
rs2.ActiveConnection = connString '连接数据库并初始化记录集


rs.Source =""&who&""
rs.Open()


do while not rs.bof or not rs.eof



    if rs.eof or rs.bof then
 
    response.write("<br>")
    response.write("没有其他订单")

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
    response.write(" 用户ID：")
    response.write(rs("uid"))
    response.write("    订单号：")
    response.write(rs("oid"))
    response.write("   商品名：")
    response.write(rs("name"))
    response.write("   单价：")
    response.write(rs("price"))
    response.write("   数量：")
    response.write(rs("amount"))
    response.write("   总价格：")
    response.write(rs("prices"))
    response.write("   付款方式：")
    response.write(rs("payment"))
    response.write("   快递单号：")
    response.write(rs("expressinfo"))
    response.write("   状态：")
    response.write(rs("situaction"))
    if not rs("situaction")="订单支付成功" then
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