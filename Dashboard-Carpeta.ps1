<#
.SYNOPSIS
    Lee todos los archivos de una carpeta y genera el dashboard de conciliación en HTML.

.DESCRIPTION
    Usa Conciliacion-Loop.ps1 (debe estar en la misma carpeta que este script). No descarga nada.
    Busca en la carpeta: Movimientos_Loop*.csv (todos), Debitos_Pendientes*.csv (el más reciente)
    y Movimientos_BankService*.csv (todos). El dashboard queda en <OutputDir>\Conciliacion_Loop_General.html
    (con filtro por fechas) y los reportes de cada día en <OutputDir>\Diario.
    Código de salida: 0 si terminó bien, 1 si hubo error.

.PARAMETER Carpeta       Carpeta con los archivos de entrada (obligatorio).
.PARAMETER LoopBasePath  CSV base de Loop para completar el período anterior (opcional).
.PARAMETER OutputDir     Carpeta de salida (por defecto, la carpeta 'salida' junto al script).

.EXAMPLE
    .\Dashboard-Carpeta.ps1 -Carpeta C:\Conciliacion\Entrada
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Carpeta,
    [string]$LoopBasePath,
    [string]$OutputDir = (Join-Path $PSScriptRoot 'salida')
)

$principal = Join-Path $PSScriptRoot 'Conciliacion-Loop.ps1'
if (-not (Test-Path -LiteralPath $principal)) {
    Write-Host "[ERROR] No se encuentra Conciliacion-Loop.ps1 en $PSScriptRoot (deben estar en la misma carpeta)." -ForegroundColor Red
    exit 1
}
& $principal @PSBoundParameters
exit $LASTEXITCODE
