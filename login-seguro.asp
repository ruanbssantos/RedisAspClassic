<!--#include file="_inc/function.asp"--> 
<!--#include file="_inc/json.asp"--> 
<%
    Response.LCID = 1046 ' REQUIRED! Set your LCID here (1046 = Brazilian). Could also be the LCID property of the page declaration or the Session.LCID property
 
    ' Verificar se o formulário foi enviado via POST
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

        ' Obter os dados do formulário
        'Dim cpf, nome, email, telefone, session_id, secretKey
        cpf = Request.Form("cpf") ' Recebe o CPF enviado pelo formulário
        nome = Request.Form("nome") ' Recebe o nome enviado pelo formulário
        email = Request.Form("email") ' Recebe o e-mail enviado pelo formulário
        telefone = Request.Form("telefone") ' Recebe o telefone enviado pelo formulário

        
        ' instantiate the class
        set JSON = New JSONobject

        JSON.Add "cpf", cpf
        JSON.Add "nome", nome
        JSON.Add "email", email
        JSON.Add "telefone", telefone

        jsonString = JSON.Serialize()

        ' Criptografar o session_id usando PowerShell
        session_id = GenerateGUID()
        'session_id = "E10F7301-3AE1-45C2-B87F-2B1145E27A2B" 
        
        ' Criptografar o session_id usando PowerShell
        'Dim encryptedSessionID
        encryptedSessionID = EncryptWithAES(session_id,secretKey) ' Chama o script PowerShell para criptografar
        
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
        call Redis.Set("session1:" & session_id, jsonString, 600) ' Cria uma associação session_id 
        call Redis.Expire("session1:" & session_id, 600) ' Define um tempo de expiração de 210 segundos

        ' Armazenar o session_id no Redis, associando-o ao CPF do usuário
        call Redis.Set("session:" & session_id, cpf, 600) ' Cria uma associação session_id 
        call Redis.Expire("session:" & session_id, 600) ' Define um tempo de expiração de 210 segundos
        
        ' Definir o cookie com o session_id criptografado via cabeçalho (somente HttpOnly)
        If Len(encryptedSessionID) > 0 Then

            call SetBase64HttpOnlyCookie("sessionid", encryptedSessionID, 60) ' Armazena o session_id criptografado no cookie por 60 minutos 
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