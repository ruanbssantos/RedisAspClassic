<!--#include file="_inc/function.asp"--> 
<%
    ' Criptografar o session_id usando PowerShell
    session_id = GenerateGUID()
    str = EncryptWithAES(session_id,secretKey) 
    Response.Write("<br/>session_id: " & session_id & "<br/><br/>")
    Response.Write("<br/>session_id criptografada: " & str)
    response.write( "<br><br><br><br>")
    
     

    call SetBase64HttpOnlyCookie("sessionid", str, 60)
    session_id_encrypted = GetBase64FromCookie("sessionid")

    response.write("<br/><br/>Valor do coockie: " & session_id_encrypted  & "<br/>")

    Response.Write("<br/>Descriptografia: " & DecryptWithAES(session_id_encrypted,secretKey)  & "<br/>")	
    'Response.Write("<br/>session_id criptografado hexa para base64: " & EncodeStringBase64(str) & "<br/>")	


    response.end
    ' Verificar se o formulário foi enviado via POST
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

        ' Obter os dados do formulário
        'Dim cpf, nome, email, telefone, session_id, secretKey
        cpf = Request.Form("cpf") ' Recebe o CPF enviado pelo formulário
        nome = Request.Form("nome") ' Recebe o nome enviado pelo formulário
        email = Request.Form("email") ' Recebe o e-mail enviado pelo formulário
        telefone = Request.Form("telefone") ' Recebe o telefone enviado pelo formulário

        ' Gerar um identificador de sessão seguro (sem caracteres especiais)
        session_id = GenerateSessionID(32) ' Gera um token de 32 caracteres
 
        ' Criptografar o session_id usando PowerShell
        'Dim encryptedSessionID
        encryptedSessionID = EncryptWithAES(session_id, secretKey) ' Chama o script PowerShell para criptografar
        
        ' Verificar o valor criptografado para garantir que não está vazio
        Response.Write("<br/>session_id criptografado: " & encryptedSessionID & "<br/>")	

        ' Conectar ao Redis usando a biblioteca redis-com-client
        Set Redis = Server.CreateObject("RedisComClient")
        'Redis.Open("192.168.15.15:6379")
        Redis.Open(vstrServer_redis)

        ' Criar uma chave única para o CPF (hash no Redis)
        Dim key
        key = "user:" & cpf

        ' Gravar os dados do usuário no Redis com TTL de 3600 segundos (60 minutos)
        call Redis.Hset(key, "nome", nome)
        call Redis.Hset(key, "email", email)
        call Redis.Hset(key, "telefone", telefone)    
        call Redis.Expire(key, 600) ' Define um tempo de expiração de 210 segundos

        ' Armazenar o session_id no Redis, associando-o ao CPF do usuário
        call Redis.Set("session:" & session_id, cpf, 600) ' Cria uma associação session_id 
        call Redis.Expire("session:" & session_id, 600) ' Define um tempo de expiração de 210 segundos
        
        ' Definir o cookie com o session_id criptografado via cabeçalho (somente HttpOnly)
        If Len(encryptedSessionID) > 0 Then
            SetHttpOnlyCookie "sessionid", encryptedSessionID, 60 ' Armazena o session_id criptografado no cookie por 60 minutos
        Else
            Response.Write("Erro: session_id criptografado está vazio!<br/>")
        End If
        
        'Response.Cookies("session_id") = encryptedSessionID ' Armazena o session_id criptografado no cookie
        'Response.Cookies("session_id").Expires = DateAdd("h", 210, Now()) ' Expira em 4 minutos
        'Response.Cookies("session_id").HttpOnly = True ' Torna o cookie inacessível via JavaScript
        'Response.Cookies("session_id").Secure = True ' Cookie será enviado apenas via HTTPS

        ' Mensagem de sucesso após o cadastro
        Response.Write("Cadastro realizado com sucesso. Sua sessão foi iniciada.<br/>")
        Response.Write("Seu session_id é: " & session_id & "<br/>")	

        ' Liberar o objeto Redis
        Set Redis = Nothing
    End If
%>
<!DOCTYPE html>
<html>
<head>
    <title>Formulário de Cadastro</title>
</head>
<body>
    <h1>Cadastro de Usuário</h1>
    <form method="post" action="">
        <label for="cpf">CPF:</label><br/>
        <input type="text" id="cpf" name="cpf" required><br/><br/>

        <label for="nome">Nome:</label><br/>
        <input type="text" id="nome" name="nome" required><br/><br/>

        <label for="email">E-mail:</label><br/>
        <input type="email" id="email" name="email" required><br/><br/>

        <label for="telefone">Telefone:</label><br/>
        <input type="text" id="telefone" name="telefone" required><br/><br/>

        <input type="submit" value="Salvar">
    </form>
</body>
</html>