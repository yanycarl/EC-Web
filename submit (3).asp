<%
if request.cookies("guango").haskeys then
else
      response.cookies("guango")("id")="0"
end if
Response.Cookies("guango")("name")="����ŵ"
response.redirect("buy.asp")%>