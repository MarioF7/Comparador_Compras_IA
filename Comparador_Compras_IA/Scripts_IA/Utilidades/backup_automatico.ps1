# Script de backup automático - Sistema Comparador Compras IA
param([string]$ProjectPath = ".")

$backupDir = Join-Path $ProjectPath "Data_Backup\Automatico\$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

# Archivos a respaldar
$filesToBackup = @(
    "Comparador_Compras_IA_Completo.xlsm",
    "Configuraciones\*.json",
    "Configuraciones\*.xml",
    "Logs\*.log"
)

foreach ($pattern in $filesToBackup) {
    $files = Get-ChildItem -Path (Join-Path $ProjectPath $pattern) -File
    foreach ($file in $files) {
        $dest = Join-Path $backupDir $file.Name
        Copy-Item $file.FullName $dest -Force
    }
}

# Comprimir backup
$zipFile = "$backupDir.zip"
Compress-Archive -Path "$backupDir\*" -DestinationPath $zipFile -Force

# Limpiar carpeta temporal
Remove-Item $backupDir -Recurse -Force

Write-Output "Backup completado: $zipFile"
