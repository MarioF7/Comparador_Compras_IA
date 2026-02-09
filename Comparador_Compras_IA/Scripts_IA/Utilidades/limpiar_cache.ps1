# Script para limpiar caché del sistema
param([string]$ProjectPath = ".")

$cacheDirs = @(
    "Cache\Imagenes",
    "Cache\Datos",
    "Cache\Temporal",
    "Temp"
)

$totalFreed = 0
foreach ($dir in $cacheDirs) {
    $fullPath = Join-Path $ProjectPath $dir
    if (Test-Path $fullPath) {
        $files = Get-ChildItem $fullPath -File -Recurse
        $size = ($files | Measure-Object -Property Length -Sum).Sum
        Remove-Item "$fullPath\*" -Recurse -Force
        $totalFreed += $size
    }
}

Write-Output "Cache limpiado: $([math]::Round($totalFreed/1MB, 2)) MB liberados"
