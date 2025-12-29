trap {
    Write-Host "`n❌ Erro inesperado: $_" -ForegroundColor Red
    Write-Host "Pressione qualquer tecla para sair..."
    pause
    exit 1
}

$TenantId     = "Seu-Tenant-ID-Aqui"
$ClientId     = "Seu-Client-ID-Aqui"
$ClientSecret = "Seu-Client-Secret-Aqui"

$TokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$GraphBaseUrl = "https://graph.microsoft.com/v1.0"

$Body = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}

Write-Host "Obtendo token..." -ForegroundColor Cyan
$TokenResponse = Invoke-RestMethod -Method Post -Uri $TokenUrl -Body $Body
$AccessToken = $TokenResponse.access_token
$Headers = @{
    "Authorization" = "Bearer $AccessToken"
    "Content-Type"  = "application/json; charset=utf-8"
}

# Coleta de dados
$NomeUsuario    = Read-Host "`n1. Nome completo do usuário"
$EmailUsuario   = Read-Host "2. E-mail do usuário"
$DataInicioStr  = Read-Host "3. Data e hora de início (ex: 2025-12-29 09:00)"
$DataFimStr     = Read-Host "4. Data e hora de fim (ex: 2025-12-30 18:00)"
$Redirecionar   = Read-Host "5. Redirecionar? (S/N)"

# Validação
try {
    $Inicio = [DateTime]::ParseExact($DataInicioStr, "yyyy-MM-dd HH:mm", $null)
    $Fim    = [DateTime]::ParseExact($DataFimStr, "yyyy-MM-dd HH:mm", $null)
    if ($Inicio -ge $Fim) { throw "Início deve ser antes do fim." }
} catch {
    Write-Host "❌ Formato inválido." -ForegroundColor Red
    pause
    exit 1
}

$TimeZone = "E. South America Standard Time"

$MensagemFinal = "Prezados(as),`n`n" +
                 "Informo que estarei de férias no período de $($Inicio.ToString("dd/MM/yyyy")) a $($Fim.ToString("dd/MM/yyyy")), retornando às minhas atividades no dia $(($Fim.AddDays(1)).ToString("dd/MM/yyyy"))."

if ($Redirecionar -match "^[sS]$") {
    $NomeRedirecionado  = Read-Host "6. Nome do substituto"
    $EmailRedirecionado = Read-Host "7. E-mail do substituto"
    $MensagemFinal += "`n`nDurante esse período, não estarei acompanhando as demandas. Para assuntos relacionados à minha área de atuação, gentileza direcionar para $NomeRedirecionado, no e-mail $EmailRedirecionado, que poderá auxiliá-los no que for necessário."
}

$MensagemFinal += "`n`nAgradeço pela compreensão.`n`nAtenciosamente,`n$NomeUsuario"

# Monta o corpo da requisição
$OofBodyObj = @{
    automaticRepliesSetting = @{
        status = "scheduled"
        scheduledStartDateTime = @{
            dateTime = $Inicio.ToString("yyyy-MM-ddTHH:mm:ss")
            timeZone = $TimeZone
        }
        scheduledEndDateTime = @{
            dateTime = $Fim.ToString("yyyy-MM-ddTHH:mm:ss")
            timeZone = $TimeZone
        }
        internalReplyMessage = $MensagemFinal
        externalReplyMessage = $MensagemFinal
    }
}

# Serializa o JSON e converte para UTF-8 EXPLÍCITO
$OofJson = $OofBodyObj | ConvertTo-Json -Depth 5
$Utf8Bytes = [System.Text.Encoding]::UTF8.GetBytes($OofJson)

Write-Host "`nEnviando configuração com codificação UTF-8 explícita..." -ForegroundColor Cyan
try {
    $OofUrl = "$GraphBaseUrl/users/$EmailUsuario/mailboxSettings"
    # Envia os bytes diretamente (não como string)
    $Response = Invoke-RestMethod -Method PATCH -Uri $OofUrl -Headers $Headers -Body $Utf8Bytes
} catch {
    Write-Host "❌ Erro: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n✅ Concluído!" -ForegroundColor Cyan
Write-Host "`nPressione qualquer tecla para sair..."
pause