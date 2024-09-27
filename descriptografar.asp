<!--#include file="_inc/function.asp"--> 
<%
    
    session_id_encrypted = GetBase64FromCookie("sessionid") 
 
    response.write "<br>Request Cookie: " & session_id_encrypted
    response.write "<br>Descriptografia: " & DecryptWithAES(session_id_encrypted,secretKey) 
    response.end

%>