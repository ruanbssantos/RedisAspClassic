<%
    vstr_scriptPath = "C:\inetpub\wwwroot\Redis\_inc\scripts"    
    ' Definir a chave secreta usada para criptografar o session_id
    secretKey = "chaveSegura123" ' Defina uma chave secreta (mínimo de 16 caracteres)
    vstrServer_redis = "127.0.0.1:6379"

    Function GenerateGUID()
        Dim objTypeLib, guid
        Set objTypeLib = Server.CreateObject("Scriptlet.TypeLib")
        guid = objTypeLib.Guid
        ' Remove as chaves {} do GUID 
        GenerateGUID = replace(replace(guid,"}",""),"{","")
        Set objTypeLib = Nothing
    End Function 

    ' Função para definir um cookie com HttpOnly via cabeçalhos
    Sub SetBase64HttpOnlyCookie(name, value, expiresInMinutes)
        Dim safeBase64Value,cookieHeader, expires

        ' Definir a data de expiração do cookie no formato GMT
        expires = Now() + (expiresInMinutes / 1440) ' Converte minutos para dias
        expires = FormatDateTime(expires, vbLongDate) & " " & Time()

        ' Montar o cabeçalho do cookie com HttpOnly (sem Secure)
        cookieHeader = name & "=" & server.URLEncode(safeBase64Value) & "; Expires=" & expires & "; Path=/; HttpOnly"
        
        ' Enviar o cabeçalho do cookie
        Response.AddHeader "Set-Cookie", cookieHeader
    End Sub 
    
    Function GetBase64FromCookie(name)
        Dim safeBase64Value, base64Value
        
        ' Recuperar o valor do cookie
        safeBase64Value = Request.Cookies(name)


        response.write("<br /> Request coockie bruto: " & safeBase64Value)

        ' Reverter os caracteres substituídos para o formato original Base64
        base64Value = Replace(safeBase64Value, "-", "+")  ' Substitui - por +
        base64Value = Replace(base64Value, "_", "/")  ' Substitui _ por /

        ' Adicionar os caracteres de preenchimento (=) novamente se necessário
        Dim paddingLength
        paddingLength = Len(base64Value) Mod 4  ' Base64 deve ser múltiplo de 4
        If paddingLength > 0 Then
            base64Value = base64Value & String(4 - paddingLength, "=")
        End If

        response.write("<br /> Request coockie tratado: " & base64Value)

        GetBase64FromCookie = base64Value
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
    
    ' Função para gerar um ID de sessão aleatório
    Function GenerateSessionID(length)
        Dim chars, i, result
        chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        result = ""

        Randomize ' Inicializa o gerador de números aleatórios

        For i = 1 To length
            result = result & Mid(chars, Int((Len(chars) * Rnd()) + 1), 1)
        Next

        GenerateSessionID = result
    End Function
%>
