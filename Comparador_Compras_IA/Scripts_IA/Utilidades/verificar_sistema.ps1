# Script de verificación del sistema
param([string]$ProjectPath = ".")

$checks = @()

# 1. Verificar archivos esenciales
$essentialFiles = @(
    "Comparador_Compras_IA_Completo.xlsm",
    "Configuraciones\config_sistema.json",
    "INSTRUCCIONES_PROYECTO.txt"
)

foreach ($file in $essentialFiles) {
    $path = Join-Path $ProjectPath $file
    $checks += @{
        Archivo = $file
        Existe = (Test-Path $path)
        Tamaño = if (Test-Path $path) { (Get-Item $path).Length } else { 0 }
    }
}

# 2. Verificar permisos
try {
    $testFile = Join-Path $ProjectPath "test_permissions.tmp"
    "test" | Out-File $testFile -Encoding UTF8
    Remove-Item $testFile -Force
    $permisos = $true
} catch {
    $permisos = $false
}

$checks += @{
    Componente = "Permisos de escritura"
    Estado = $permisos
}

# 3. Verificar espacio
$drive = (Get-PSDrive -Name $env:SystemDrive[0])
$checks += @{
    Componente = "Espacio en disco"
    Estado = ($drive.Free -gt 100MB)
    Libre = "$([math]::Round($drive.Free/1MB, 2)) MB"
}

# Mostrar resultados
$checks | ForEach-Object {
    $status = if ($_.Estado -or ($_.Existe -eq $true)) { "OK" } else { "ERROR" }
    Write-Host "[$status] $($_.Archivo ?? $_.Componente)" -ForegroundColor $(if ($status -eq "OK") { "Green" } else { "Red" })
}
