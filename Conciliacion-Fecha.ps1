<#
.SYNOPSIS
    Baja de Sybase los datos de una fecha (o un rango) que se indica por parámetro y corre la conciliación.

.DESCRIPTION
    Usa Conciliacion-Loop.ps1 (debe estar en la misma carpeta). Débitos pendientes: se bajan todos.
    Movimientos de Loop: los del día -Fecha, o del rango -Fecha a -FechaHasta. No aplica la regla de
    feriados ni de fines de semana: corre siempre el día indicado.
    Código de salida: 0 si terminó bien, 1 si hubo error.

.PARAMETER Fecha               Día a bajar y conciliar (dd/MM/yyyy o yyyy-MM-dd). Obligatorio.
.PARAMETER FechaHasta          Último día del rango (por defecto, el mismo día).
.PARAMETER TransferenciasPath  CSV de transferencias QNET (por defecto .\Movimientos_BankService.csv).
.PARAMETER LoopBasePath        CSV base de Loop para completar el período anterior (opcional).
.PARAMETER Dsn                 DSN ODBC de Sybase.
.PARAMETER Server              Servidor ASE, si no se usa -Dsn ni -ConnectionString.
.PARAMETER Port                Puerto ASE (por defecto 5000).
.PARAMETER Driver              Driver ODBC (por defecto 'Adaptive Server Enterprise').
.PARAMETER ConnectionString    Cadena de conexión ODBC completa.
.PARAMETER Usuario             Usuario de Sybase; la clave se pide en una ventana emergente.
.PARAMETER Credencial          Credencial ya armada (evita la ventana).
.PARAMETER Timeout             Segundos máximos por consulta (por defecto 300).
.PARAMETER DescargaDir         Carpeta de los CSV bajados (por defecto <OutputDir>\Entrada).
.PARAMETER IncluirPeriodo      Genera además el Excel y el HTML de todo el período.
.PARAMETER OutputDir           Carpeta de salida (por defecto, la carpeta 'salida' junto al script).

.EXAMPLE
    .\Conciliacion-Fecha.ps1 -Fecha 26/10/2026 -Server NOMBRE_O_IP -Usuario TU_USUARIO `
        -TransferenciasPath .\Movimientos_BankService.csv -OutputDir .\salida

.EXAMPLE
    .\Conciliacion-Fecha.ps1 -Fecha 20/10/2026 -FechaHasta 23/10/2026 -Dsn BANKSYS `
        -TransferenciasPath .\Movimientos_BankService.csv
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Fecha,
    [string]$FechaHasta,
    [string]$TransferenciasPath = (Join-Path (Get-Location).Path 'Movimientos_BankService.csv'),
    [string]$LoopBasePath,
    [string]$Dsn,
    [string]$ConnectionString,
    [string]$Server,
    [int]$Port = 5000,
    [string]$Driver = 'Adaptive Server Enterprise',
    [string]$Usuario,
    [System.Management.Automation.PSCredential]$Credencial,
    [int]$Timeout = 300,
    [string]$DescargaDir,
    [switch]$IncluirPeriodo,
    [string]$OutputDir = (Join-Path $PSScriptRoot 'salida')
)

$principal = Join-Path $PSScriptRoot 'Conciliacion-Loop.ps1'
if (-not (Test-Path -LiteralPath $principal)) {
    Write-Host "[ERROR] No se encuentra Conciliacion-Loop.ps1 en $PSScriptRoot (deben estar en la misma carpeta)." -ForegroundColor Red
    exit 1
}
& $principal @PSBoundParameters
exit $LASTEXITCODE
