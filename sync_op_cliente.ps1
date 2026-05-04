# ===========================================
# CONFIGURAÇÃO
# ===========================================
$SqlServer   = "172.16.2.9,1433"
$Database    = "SBO_BALDI_PRD"
$SqlUser     = "baldi"
$SqlPassword = "b4ldi@2023"

# URL do Web App publicada no Apps Script
$WebAppUrl   = "https://script.google.com/macros/s/AKfycbwD0AsRrOhWY4oua07E6HOKOtBCi8_xnJcwexE8W4r798K5p2C6_9qIYEDk4q5IT4w/exec"

# Token igual ao expectedToken do Apps Script
$ApiToken    = "true"

# Query final informada por você
$Query = @"
SELECT
    CAST(belnr_id AS VARCHAR(50))  AS op,
    CAST(kndname  AS VARCHAR(255)) AS cliente
FROM beas_fthaupt
WHERE abgkz='n'
"@

# ===========================================
# EXECUÇÃO
# ===========================================
try {
    Add-Type -AssemblyName "System.Data"

    $connString = "Server=$SqlServer;Database=$Database;User ID=$SqlUser;Password=$SqlPassword;Encrypt=False;TrustServerCertificate=True;Connection Timeout=60;"
    $conn = New-Object System.Data.SqlClient.SqlConnection($connString)
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandTimeout = 180
    $cmd.CommandText = $Query

    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $table = New-Object System.Data.DataTable
    [void]$adapter.Fill($table)

    $rows = @()
    foreach ($r in $table.Rows) {
        $op = [string]$r["op"]
        $cliente = [string]$r["cliente"]

        if (-not [string]::IsNullOrWhiteSpace($op)) {
            $rows += @{
                op      = $op.Trim()
                cliente = $cliente.Trim()
            }
        }
    }

    $bodyObj = @{
        token = $ApiToken
        rows  = $rows
    }

    $jsonBody = $bodyObj | ConvertTo-Json -Depth 8

    $response = Invoke-RestMethod `
        -Method POST `
        -Uri $WebAppUrl `
        -ContentType "application/json; charset=utf-8" `
        -Body $jsonBody `
        -TimeoutSec 300

    Write-Host "Sincronização concluída."
    Write-Host ("Linhas enviadas: " + $rows.Count)
    Write-Host ("Resposta: " + ($response | ConvertTo-Json -Depth 8))
}
catch {
    Write-Error ("Falha: " + $_.Exception.Message)
}
finally {
    if ($conn -and $conn.State -eq 'Open') {
        $conn.Close()
    }
}