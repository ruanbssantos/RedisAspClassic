param (
    [string]$Text,        # Texto a ser criptografado
    [string]$Password     # Senha para a chave AES
)

try {
    # Gerar uma chave AES a partir da senha
    $Salt = [System.Text.Encoding]::UTF8.GetBytes("SaltValue")
    $Rfc2898DeriveBytes = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $Salt, 10000)

    $Key = $Rfc2898DeriveBytes.GetBytes(32)  # Chave de 256 bits (32 bytes)
    $IV = $Rfc2898DeriveBytes.GetBytes(16)   # Vetor de inicialização (16 bytes)

    # Converter o texto para bytes
    $Data = [System.Text.Encoding]::UTF8.GetBytes($Text)

    # Configurar o AES para criptografar
    $AES = New-Object System.Security.Cryptography.AesManaged
    $AES.Key = $Key
    $AES.IV = $IV
    $AES.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $AES.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7

    $Encryptor = $AES.CreateEncryptor()
    $EncryptedData = $Encryptor.TransformFinalBlock($Data, 0, $Data.Length)

    # Converter para Base64 para facilitar a leitura e transporte
    $EncryptedString = [Convert]::ToBase64String($EncryptedData)
    $EncryptedString

    # Substituir o caractere "+" por "%2B" para evitar problemas em URLs
    #$SafeBase64String = $EncryptedBase64String -replace "\+", "%2B"

    # Retornar a string Base64 segura
    $SafeBase64String
} catch {
    Write-Host "Erro durante a criptografia: $_"
}
