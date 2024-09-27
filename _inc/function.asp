<%
    vstr_scriptPath = "C:\inetpub\wwwroot\RedisAspClassic\_inc\scripts"    
    ' Definir a chave secreta usada para criptografar o session_id
    secretKey = "0A490C2A-33C7-4FD9-9E5F-31D3E6EC7F68" ' Defina uma chave secreta (mínimo de 16 caracteres)
    vstrServer_redis = "127.0.0.1:6379"

    Function GenerateGUID()
        Dim objTypeLib, guid
        Set objTypeLib = Server.CreateObject("Scriptlet.TypeLib")
        guid = objTypeLib.Guid
        ' Remove as chaves {} do GUID 

        'NÃO PODE SER MAIOR QUE 32 POIS DA ERRO NA DESCRIPTOGRAFIA
        vstr = left(replace(replace(replace(guid,"}",""),"{",""),"-",""),32)
        response.write len(vstr) & "<br />"
        
        GenerateGUID = vstr
        Set objTypeLib = Nothing
    End Function 

    ' Função para definir um cookie com HttpOnly via cabeçalhos
    Sub SetBase64HttpOnlyCookie(name, value, expiresInMinutes)
        Dim expires
 
        ' Definir a data de expiração do cookie no formato GMT
        expires = Now() + (expiresInMinutes / 1440) ' Converte minutos para dias
        expires = FormatDateTime(expires, vbLongDate) & " " & Time()

        ' Montar o cabeçalho do cookie com HttpOnly (sem Secure)
        cookieHeader = name & "=" & Server.URLEncode(value) & "; Expires=" & expires & "; Path=/; HttpOnly"
        
        ' Enviar o cabeçalho do cookie
        Response.AddHeader "Set-Cookie", cookieHeader
    End Sub 
    
    Function GetBase64FromCookie(name) 

        ' Recuperar o valor do cookie
        vstr_value = Request.Cookies(name)
        vstr_value = URLDecode(vstr_value)

        ' Adicionar os caracteres de preenchimento (=) novamente se necessário
        Dim paddingLength
        paddingLength = Len(vstr_value) Mod 4  ' Base64 deve ser múltiplo de 4
        If paddingLength > 0 Then
            vstr_value = vstr_value & String(4 - paddingLength, "=")
        End If

        GetBase64FromCookie = vstr_value
    End Function

    Function URLDecode(ByVal str)
        ' Substitui %xx pela letra correspondente (xx é o código hexadecimal)
        Dim i, hexChar
        URLDecode = str
        i = InStr(URLDecode, "%")
        
        Do While i > 0
            ' Extrai o valor hexadecimal
            hexChar = Mid(URLDecode, i + 1, 2)
            
            ' Substitui a sequência %xx pelo caractere correspondente
            URLDecode = Left(URLDecode, i - 1) & Chr("&H" & hexChar) & Mid(URLDecode, i + 3)
            
            ' Procura o próximo % na string
            i = InStr(URLDecode, "%")
        Loop
        
        ' Substitui os sinais de adição por espaços (já que + representa um espaço em URLs codificadas)
        URLDecode = URLDecode
    End Function


    ' Função para remover espaços em branco e quebras de linha extras
    Function CleanString(str)
        ' Remover espaços extras no início e fim
        str = Trim(str)
        
        ' Remover todas as quebras de linha
        str = Replace(str, vbCrLf, "")
        str = Replace(str, vbLf, "")
        str = Replace(str, vbCr, "")
        
        ' Retornar a string limpa
        CleanString = str
    End Function

    ' Função para criptografar texto usando PowerShell AES
    Function EncryptWithAES(text, key)
        Dim shell, command, execObject, stdOut, stdErr
        Set shell = CreateObject("WScript.Shell")

        ' Modificar o comando para usar o cmd.exe para chamar o PowerShell
        command = "cmd.exe /c powershell.exe -ExecutionPolicy Bypass -File " & vstr_scriptPath & "\Encrypt-StringAES.ps1 -Text """ & text & """ -Password """ & key & """"
        
        ' Tentar executar o comando via cmd.exe
        On Error Resume Next
        Set execObject = shell.Exec(command)

        If Err.Number <> 0 Then
            response.write "Erro ao chamar o PowerShell via cmd: " & Err.Description & "<br/>"
            EncryptWithAES = empty
        Else
            stdOut = execObject.StdOut.ReadAll() ' Ler a saída do comando
            stdErr = execObject.StdErr.ReadAll() ' Ler qualquer erro
            If Len(stdErr) > 0 Then     
                response.write "<br /><br />Erro do PowerShell via cmd: " & stdErr & "<br/>"
                EncryptWithAES = empty
            Else
                EncryptWithAES = CleanString(stdOut) ' Retorna o texto criptografado
            End If
        End If

        Set shell = Nothing ' Libera o objeto WScript.Shell
    End Function

    ' Função para descriptografar texto usando PowerShell AES
    Function DecryptWithAES(encryptedText, key)
        Dim shell, command, execObject, stdOut, stdErr
        Set shell = CreateObject("WScript.Shell")

        ' Comando para executar o script PowerShell para descriptografar o texto
        command = "cmd.exe /c powershell.exe -ExecutionPolicy Bypass -File " & vstr_scriptPath & "\Decrypt-StringAES.ps1 -EncryptedText """ & encryptedText & """ -Password """ & key & """"

        'response.write command
        
        ' Tentar executar o comando via cmd.exe
        On Error Resume Next
        Set execObject = shell.Exec(command)

        If Err.Number <> 0 Then
            DecryptWithAES = empty
            response.write "Erro ao chamar o PowerShell para descriptografar: " & Err.Description & "<br/>"
        Else
            stdOut = execObject.StdOut.ReadAll() ' Ler a saída do comando
            stdErr = execObject.StdErr.ReadAll() ' Ler qualquer erro
            If Len(stdErr) > 0 Then
                response.write "Erro do PowerShell (descriptografar): " & stdErr & "<br/>"
                DecryptWithAES = empty
            Else
                DecryptWithAES = CleanString(stdOut) ' Retorna o texto descriptografado
            End If
        End If

        Set shell = Nothing ' Libera o objeto WScript.Shell
    End Function


    Function EncodeStringBase64(sText)
        Dim oXML, oNode
        Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
        Set oNode = oXML.CreateElement("base64")
        oNode.dataType = "bin.base64"
        oNode.nodeTypedValue =Stream_StringToBinary(sText)
        EncodeStringBase64 = oNode.text
        Set oNode = Nothing
        Set oXML = Nothing
    End Function

    Function Stream_StringToBinary(Text)
        Const adTypeText = 2
        Const adTypeBinary = 1

        'Create Stream object
        Dim BinaryStream 'As New Stream
        Set BinaryStream = CreateObject("ADODB.Stream")

        'Specify stream type - we want To save text/string data.
        BinaryStream.Type = adTypeText

        'Specify charset For the source text (unicode) data.
        BinaryStream.CharSet = "us-ascii"

        'Open the stream or write text/string data To the object
        BinaryStream.Open
        BinaryStream.WriteText Text

        'Change stream type To binary
        BinaryStream.Position = 0
        BinaryStream.Type = adTypeBinary

        'Ignore first two bytes - sign of
        BinaryStream.Position = 0

        'Open the stream or get binary data from the object
        Stream_StringToBinary = BinaryStream.Read

        Set BinaryStream = Nothing
    End Function

    Function DecodeStringBase64(ByVal vCode)
        Dim oXML, oNode
        Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
        Set oNode = oXML.CreateElement("base64")
        oNode.dataType = "bin.base64"
        oNode.text = vCode
        DecodeStringBase64 = Stream_BinaryToString(oNode.nodeTypedValue)
        Set oNode = Nothing
        Set oXML = Nothing
    End Function

    Private Function Stream_BinaryToString(Binary)
        Const adTypeText = 2
        Const adTypeBinary = 1
        Dim BinaryStream 'As New Stream
        Set BinaryStream = CreateObject("ADODB.Stream")
        BinaryStream.Type = adTypeBinary
        BinaryStream.Open
        BinaryStream.Write Binary
        BinaryStream.Position = 0
        BinaryStream.Type = adTypeText
        BinaryStream.CharSet = "us-ascii"
        Stream_BinaryToString = BinaryStream.ReadText
        Set BinaryStream = Nothing
    End Function 
%>
