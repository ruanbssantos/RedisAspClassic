param (
    [string]$EncryptedText,  # Texto criptografado em Base64
    [string]$Password        # Senha para a chave AES
)

try {
    # Gerar uma chave AES a partir da senha
    $Salt = [System.Text.Encoding]::UTF8.GetBytes("SaltValue")
    $Rfc2898DeriveBytes = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $Salt, 10000)

    $Key = $Rfc2898DeriveBytes.GetBytes(32)  # Chave de 256 bits (32 bytes)
    $IV = $Rfc2898DeriveBytes.GetBytes(16)   # Vetor de inicialização (16 bytes)

    # Converter o texto criptografado de Base64 para bytes
    $EncryptedData = [Convert]::FromBase64String($EncryptedText)

    # Configurar o AES para descriptografar
    $AES = New-Object System.Security.Cryptography.AesManaged
    $AES.Key = $Key
    $AES.IV = $IV
    $AES.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $AES.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7

    $Decryptor = $AES.CreateDecryptor()
    $DecryptedData = $Decryptor.TransformFinalBlock($EncryptedData, 0, $EncryptedData.Length)

    # Converter os bytes descriptografados de volta para uma string
    $DecryptedString = [System.Text.Encoding]::UTF8.GetString($DecryptedData)

    $DecryptedString
} catch {
    Write-Host "Erro durante a descriptografia: $_"
}
