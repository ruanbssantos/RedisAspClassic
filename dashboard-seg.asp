<!--#include file="_inc/function.asp"--> 
<!--#include file="_inc/json.asp"--> 

<% 
    Response.LCID = 1046

    ' instantiate the class
    set sessionJson = New JSONobject

    ' Conectar ao Redis usando a biblioteca redis-com-client
    Set Redis = Server.CreateObject("RedisComClient")
    Redis.Open(vstrServer_redis)

    session_id_encrypted = Request.Cookies("sessionid")

    Response.Write("<br/>session_id criptografado: " & session_id_encrypted & "<br/>")

    If Not IsNull(session_id_encrypted) And session_id_encrypted <> "" Then
        ' Descriptografar o session_id
        session_id = DecryptWithAES(session_id_encrypted, secretKey)
        
        Response.Write("<br/>session_id Descriptografado: |" & session_id & "|<br/>")
        
        If Len(session_id) > 0 and redis.Exists("session:" & CleanString(session_id)) Then
 
            ' Recuperar o CPF associado ao session_id
            cpf = Redis.Get("session:" & CleanString(session_id))
            chaveHashe = "user:" & cpf

            sessionJson.Parse(Redis.Get("session1:" & CleanString(session_id)))

            

            Response.Write("<br/>session: " & session_id & "<br/>")
            Response.Write("<br/>CPF: " & cpf & "<br/>")
            Response.Write("<br/>tempo:" & redis.TTL(chaveHashe) & "<br/>")
            Response.Write("<br/>Campos em um hash:" & redis.Hlen(chaveHashe) & "<br/>")
            Response.Write("<br/>Existe:" & redis.Exists(chaveHashe) & "<br/>")
            
            If redis.Exists(chaveHashe) Then 
 
                nome = Redis.Hget(chaveHashe, "nome")
                email = Redis.Hget(chaveHashe, "email")
                telefone = Redis.Hget(chaveHashe, "telefone")
            Else
                Response.Write "<p>Sessão expirada ou inválida. Faça login novamente.</p>"
                Response.End()
            End If
        Else
            Response.Write "<p>Erro na descriptografia do session_id. Faça login novamente.</p>"
            Response.End()
        End If
    Else
        Response.Write "<p>Nenhuma sessão ativa encontrada. Faça login.</p>"
        Response.End()
    End If

    ' Liberar o objeto Redis
    Set Redis = Nothing
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard</title>
    <style>
        body {
            margin: 0;
            font-family: Arial, sans-serif;
        }
        .top-bar {
            background-color: #4CAF50;
            color: white;
            padding: 10px;
            text-align: center;
        }
        .side-menu {
            height: 100%;
            width: 200px;
            position: fixed;
            top: 0;
            left: 0;
            background-color: #333;
            padding-top: 20px;
        }
        .side-menu a {
            padding: 10px 15px;
            text-decoration: none;
            color: white;
            display: block;
        }
        .side-menu a:hover {
            background-color: #575757;
        }
        .content {
            margin-left: 200px;
            padding: 20px;
        }
        .footer {
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background-color: #4CAF50;
            color: white;
            text-align: center;
            padding: 10px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
        }
        table, th, td {
            border: 1px solid black;
        }
        th, td {
            padding: 10px;
            text-align: left;
        }
        th {
            background-color: #4CAF50;
            color: white;
        }
    </style>
</head>
<body>

    <!-- Barra superior -->
    <div class="top-bar">
        <h1>Dashboard - Bem-vindo(a), <%=nome%>!</h1>
    </div> 
    <!-- Conteúdo principal -->
    <div class="content">
        <h2>Dados do Usuario</h2>

        <!-- Tabela de dados -->
        <table>
            <tr>
                <th colspan="100%">Valor unitario</th>
            </tr>
            <tr>
                <th>Nome</th>
                <th>E-mail</th>
                <th>Telefone</th>
            </tr>
            <tr>
                <td><%=nome%></td>
                <td><%=email%></td>
                <td><%=telefone%></td>
            </tr>
        </table>

         <table>
            <tr>
                <tr>
                    <th colspan="100%">Sessao em json</th>
                </tr>
            </tr>
            <tr>
                <td><%=sessionJson.value("nome")%></td>
                <td><%=sessionJson.value("email")%></td>
                <td><%=sessionJson.value("telefone")%></td>
            </tr>
        </table>
    </div>

    <!-- Barra inferior -->
    <div class="footer">
        <p>© 2024 Sua Empresa - Todos os direitos reservados</p>
    </div>

</body>
</html>