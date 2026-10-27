<#
.SYNOPSIS
    Todo en uno: baja de Sybase los débitos pendientes y los movimientos de Loop (opcional, con
    -Descargar) y después concilia las transferencias QNET contra Loop y los débitos pendientes.
    Genera un Excel (o CSV de respaldo) y un dashboard HTML.

.DESCRIPTION
    Compatible con Windows PowerShell 5.1 y PowerShell 7.
    Entradas: Loop (partidas), Débitos pendientes y Transferencias QNET (CSV).
    Salidas: un Excel y un HTML por día en <OutputDir>\Diario (Conciliacion_Loop_<yyyyMMdd>.xlsx/.html)
    y un dashboard general (Conciliacion_Loop_General.html) con filtro por fechas, que se arma
    leyendo los datos embebidos en los HTML diarios de esa carpeta.
    Código de salida: 0 si no hubo excepciones de ejecución, 1 si las hubo.
    No se vuelcan nombres, CUIT ni CBU en las salidas.

.PARAMETER LoopPath            CSV de Loop (obligatorio).
.PARAMETER PendientesPath      CSV de débitos pendientes (obligatorio).
.PARAMETER TransferenciasPath  CSV de transferencias QNET (obligatorio).
.PARAMETER LoopBasePath        CSV base de Loop para completar el período anterior (opcional).
.PARAMETER IncluirPeriodo      Genera además el Excel y el HTML de todo el período.
.PARAMETER SoloGeneral         Solo reconstruye el dashboard general desde los HTML diarios existentes.
.PARAMETER OutputDir           Carpeta de salida (por defecto, la carpeta actual).
.PARAMETER Descargar           Paso 1: baja Debitos_Pendientes.csv y Movimientos_Loop.csv de Sybase y los usa
                               como -PendientesPath y -LoopPath (no se pasan esos dos parámetros).
.PARAMETER Dsn                 (con -Descargar) DSN ODBC de Sybase.
.PARAMETER Server              (con -Descargar) Servidor ASE, si no se usa -Dsn ni -ConnectionString.
.PARAMETER Port                (con -Descargar) Puerto ASE (por defecto 5000).
.PARAMETER Driver              (con -Descargar) Driver ODBC (por defecto 'Adaptive Server Enterprise').
.PARAMETER ConnectionString    (con -Descargar) Cadena de conexión ODBC completa.
.PARAMETER Usuario             (con -Descargar) Usuario de Sybase; la clave se pide en una ventana emergente.
.PARAMETER Credencial          (con -Descargar) Credencial ya armada (evita la ventana).
.PARAMETER Fecha               (con -Descargar) Día puntual a bajar (por defecto, desde el último día hábil anterior; lunes: desde el viernes).
.PARAMETER Timeout             (con -Descargar) Segundos máximos por consulta (por defecto 300).
.PARAMETER DescargaDir         (con -Descargar) Carpeta de los CSV bajados (por defecto <OutputDir>\Entrada).

.EXAMPLE
    .\Conciliacion-Loop.ps1 -LoopPath .\SQL14.csv -LoopBasePath .\Tr_Loop_116261.csv `
        -PendientesPath .\SQL15.csv `
        -TransferenciasPath .\transferencias-qnet-2026-10-01_2026-10-23.csv `
        -OutputDir .\salida

.EXAMPLE
    # Todo en uno: baja de Sybase y concilia (sábados y domingos no se ejecuta)
    .\Conciliacion-Loop.ps1 -Descargar -Server NOMBRE_O_IP -Usuario TU_USUARIO `
        -TransferenciasPath .\Movimientos_BankService.csv -OutputDir .\salida
#>
#Requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'Conciliar')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Conciliar')][string]$LoopPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Conciliar')][string]$PendientesPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Conciliar')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Completo')][string]$TransferenciasPath,
    [Parameter(ParameterSetName = 'Conciliar')]
    [Parameter(ParameterSetName = 'Completo')][string]$LoopBasePath,
    [Parameter(ParameterSetName = 'Conciliar')]
    [Parameter(ParameterSetName = 'Completo')][switch]$IncluirPeriodo,
    [Parameter(Mandatory = $true, ParameterSetName = 'Completo')][switch]$Descargar,
    [Parameter(ParameterSetName = 'Completo')][string]$Dsn,
    [Parameter(ParameterSetName = 'Completo')][string]$ConnectionString,
    [Parameter(ParameterSetName = 'Completo')][string]$Server,
    [Parameter(ParameterSetName = 'Completo')][int]$Port = 5000,
    [Parameter(ParameterSetName = 'Completo')][string]$Driver = 'Adaptive Server Enterprise',
    [Parameter(ParameterSetName = 'Completo')][string]$Usuario,
    [Parameter(ParameterSetName = 'Completo')][System.Management.Automation.PSCredential]$Credencial,
    [Parameter(ParameterSetName = 'Completo')][datetime]$Fecha,
    [Parameter(ParameterSetName = 'Completo')][int]$Timeout = 300,
    [Parameter(ParameterSetName = 'Completo')][string]$DescargaDir,
    [Parameter(Mandatory = $true, ParameterSetName = 'General')][switch]$SoloGeneral,
    [string]$OutputDir = (Get-Location).Path
)

$ErrorActionPreference = 'Stop'

# Acentos: la consola y el archivo en UTF-8 (el script se guarda con BOM para Windows PowerShell 5.1)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# ---------------------------------------------------------------------------
# Constantes y formato
# ---------------------------------------------------------------------------
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture
$script:Tol = [decimal]0.005
# Feriados y días no laborables 2026: la descarga no se ejecuta esos días (ni sábados y domingos)
$script:Feriados = @{
    '2026-01-01' = 'Año nuevo'; '2026-02-16' = 'Carnaval'; '2026-02-17' = 'Carnaval'
    '2026-03-23' = 'Día no laborable con fines turísticos'; '2026-03-24' = 'Día Nacional de la Memoria por la Verdad y la Justicia'
    '2026-04-02' = 'Día del Veterano y de los Caídos en la Guerra de Malvinas y Jueves Santo'; '2026-04-03' = 'Viernes Santo'
    '2026-05-01' = 'Día del Trabajador'; '2026-05-25' = 'Día de la Revolución de Mayo'
    '2026-06-15' = 'Paso a la Inmortalidad del Gral. Güemes'; '2026-07-09' = 'Día de la Independencia'
    '2026-07-10' = 'Día no laborable con fines turísticos'; '2026-08-17' = 'Paso a la Inmortalidad del Gral. San Martín'
    '2026-10-12' = 'Día de la Raza'; '2026-11-09' = 'Feriado nacional por la visita de Su Santidad el Papa León XIV'
    '2026-11-10' = 'Feriado bancario por la visita de Su Santidad el Papa León XIV'; '2026-11-23' = 'Día de la Soberanía Nacional'
    '2026-12-07' = 'Día no laborable con fines turísticos'; '2026-12-08' = 'Día de la Inmaculada Concepción de María'
    '2026-12-25' = 'Navidad'
}
$script:NfAR = New-Object System.Globalization.NumberFormatInfo
$script:NfAR.NumberDecimalSeparator = ','
$script:NfAR.NumberGroupSeparator = '.'
$script:NfAR.NumberGroupSizes = @(3)
$script:NfAR.NegativeSign = '-'
$script:OrdenTipos = @(
    'Más de 2 partidas', 'Una partida', 'Importes distintos', 'Pendiente con importe distinto',
    'Pendiente sin asiento en Loop', 'Pendiente sobre comprobante completo',
    'Pendiente: importe distinto a transfer', 'Pendiente: cuenta distinta a transfer',
    'FAILED contabilizado', 'APPROVED sin registro', 'Sin operación en transfers',
    'Importe distinto a transfer', 'Moneda distinta a transfer', 'Cuenta distinta a transfer',
    'CBU origen no corresponde a la moneda', 'Débito pendiente'
)

function Write-Log {
    param([string]$Mensaje, [string]$Nivel = 'INFO')
    $color = 'Gray'
    if ($Nivel -eq 'WARN') { $color = 'Yellow' }
    elseif ($Nivel -eq 'ERROR') { $color = 'Red' }
    elseif ($Nivel -eq 'OK') { $color = 'Green' }
    Write-Host ('[{0}] {1}' -f $Nivel, $Mensaje) -ForegroundColor $color
}

function Format-Num {
    param([decimal]$Valor)
    return $Valor.ToString('N2', $script:NfAR)
}

function Format-Periodo {
    param($Desde, $Hasta)
    if ($Desde -eq $Hasta) { return (Format-Fecha $Desde) }
    return ('{0} al {1}' -f (Format-Fecha $Desde), (Format-Fecha $Hasta))
}

function Format-Importe {
    param([decimal]$Valor, [int]$MonCod)
    if ($MonCod -eq 2) { return 'US$ ' + (Format-Num $Valor) }
    return '$ ' + (Format-Num $Valor)
}

function Format-Fecha {
    param($Fecha)
    if ($null -eq $Fecha) { return '' }
    return ([datetime]$Fecha).ToString('dd/MM/yyyy', $script:Inv)
}

function Get-NombreMoneda {
    param([int]$MonCod)
    if ($MonCod -eq 0) { return 'Pesos' }
    if ($MonCod -eq 2) { return 'USD' }
    return "Moneda $MonCod"
}

function Get-Enc { param([string]$Texto) return [System.Net.WebUtility]::HtmlEncode($Texto) }

# ---------------------------------------------------------------------------
# Conversión y lectura tolerante
# ---------------------------------------------------------------------------
function Convert-Decimal {
    param($Texto)
    if ($null -eq $Texto) { return $null }
    $s = ([string]$Texto).Trim().Trim('"')
    if ($s -eq '') { return $null }
    $d = [decimal]0
    if ([decimal]::TryParse($s, [System.Globalization.NumberStyles]::Float, $script:Inv, [ref]$d)) { return $d }
    return $null
}

function Convert-Entero {
    param($Texto)
    if ($null -eq $Texto) { return $null }
    $s = ([string]$Texto).Trim().Trim('"')
    if ($s -eq '') { return $null }
    $n = [long]0
    if ([long]::TryParse($s, [System.Globalization.NumberStyles]::AllowLeadingSign, $script:Inv, [ref]$n)) { return $n }
    return $null
}

function Convert-Fecha {
    param($Texto, [string[]]$Formatos)
    if ($null -eq $Texto) { return $null }
    $s = ([string]$Texto).Trim().Trim('"')
    if ($s -eq '') { return $null }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, $Formatos, $script:Inv, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt }
    return $null
}

function Format-Comnro {
    # Normaliza el número de comprobante (sin espacios ni ceros a la izquierda)
    param($Texto)
    $s = ([string]$Texto).Trim().Trim('"')
    if ($s -match '^\d+$') {
        $s = $s.TrimStart('0')
        if ($s -eq '') { $s = '0' }
    }
    return $s
}

function Get-Prop {
    param($Fila, [string]$Nombre)
    $p = $Fila.PSObject.Properties[$Nombre]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Import-CsvTolerante {
    # Lee un CSV tolerando CRLF/LF mezclados, última línea sin salto y BOM.
    param([string]$Path, [switch]$CortarEnBlanco)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "No existe el archivo: $Path" }
    $texto = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($texto)) { throw "El archivo está vacío: $Path" }
    $lineas = $texto -split "`r`n|`n|`r"
    $i = 0
    while ($i -lt $lineas.Count -and $lineas[$i].Trim().Trim([char]0xFEFF) -eq '') { $i++ }
    $out = New-Object 'System.Collections.Generic.List[string]'
    for (; $i -lt $lineas.Count; $i++) {
        if ($lineas[$i].Trim() -eq '') {
            if ($CortarEnBlanco) { break }
            continue
        }
        $out.Add($lineas[$i])
    }
    if ($out.Count -eq 0) { throw "El archivo no tiene datos: $Path" }
    $out[0] = $out[0].TrimStart([char]0xFEFF)
    $delim = ','
    $max = -1
    foreach ($c in @(',', ';', "`t")) {
        $n = [regex]::Matches($out[0], [regex]::Escape($c)).Count
        if ($n -gt $max) { $max = $n; $delim = $c }
    }
    $header = @($out[0] -split [regex]::Escape($delim) | ForEach-Object { $_.Trim().Trim('"') })
    $rows = @($out | ConvertFrom-Csv -Delimiter $delim)
    return [pscustomobject]@{ Header = $header; Rows = $rows; Archivo = (Split-Path -Leaf $Path) }
}

function Test-Columnas {
    param([string[]]$Header, [string[]]$Requeridas, [string]$Archivo)
    $faltan = @($Requeridas | Where-Object { $Header -notcontains $_ })
    if ($faltan.Count -gt 0) { throw "Faltan columnas en ${Archivo}: $($faltan -join ', ')" }
}

function Write-Malformadas {
    param($Lista, [string]$Archivo)
    if ($Lista.Count -eq 0) { return }
    Write-Log "$Archivo : $($Lista.Count) fila(s) mal formada(s) descartada(s)." 'WARN'
    $Lista | Select-Object -First 5 | ForEach-Object { Write-Log "  $_" 'WARN' }
}

# ---------------------------------------------------------------------------
# Lectura de archivos
# ---------------------------------------------------------------------------
function Get-FirmaLoop {
    param($Sist, $Mon, $Cue, $Fecha, $Comnro, $Mov, $Signo, $Acc, $Imp)
    return ('{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}' -f $Sist, $Mon, $Cue, $Fecha.ToString('yyyyMMdd', $script:Inv), $Comnro, $Mov, $Signo,
        $Acc.ToString('0.00########', $script:Inv), $Imp.ToString('0.00########', $script:Inv))
}

function Read-Loop {
    param([string]$Path)
    $csv = Import-CsvTolerante -Path $Path -CortarEnBlanco
    $req = @('sistcod', 'moncod', 'cuecod', 'Fecha_Mov', 'comnro', 'movcod', 'signo', 'accimp', 'Importe')
    Test-Columnas $csv.Header $req $csv.Archivo
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $malas = New-Object 'System.Collections.Generic.List[string]'
    $n = 1
    foreach ($r in $csv.Rows) {
        $n++
        try {
            $sist = Convert-Entero $r.sistcod
            $mon = Convert-Entero $r.moncod
            $cue = Convert-Entero $r.cuecod
            $mov = Convert-Entero $r.movcod
            $signo = Convert-Entero $r.signo
            $fecha = Convert-Fecha $r.Fecha_Mov @('dd/MM/yyyy', 'd/M/yyyy')
            $acc = Convert-Decimal $r.accimp
            $imp = Convert-Decimal $r.Importe
            $com = Format-Comnro $r.comnro
            if ($null -eq $sist -or $null -eq $mon -or $null -eq $mov -or $null -eq $signo) { throw 'sistcod/moncod/movcod/signo inválido' }
            if ($null -eq $fecha) { throw "Fecha_Mov inválida '$($r.Fecha_Mov)'" }
            if ($com -eq '') { throw 'comnro vacío' }
            if ($null -eq $imp -and $null -eq $acc) { throw 'sin importe' }
            if ($null -eq $imp) { $imp = $signo * $acc }
            if ($null -eq $acc) { $acc = [Math]::Abs($imp) }
            if ($null -eq $cue) { $cue = [long]0 }
            $rows.Add([pscustomobject]@{
                    Sist = [int]$sist; Mon = [int]$mon; Cue = [long]$cue; Fecha = $fecha.Date; Comnro = $com
                    Mov = [int]$mov; Signo = [int]$signo; Acc = [decimal]$acc; Imp = [decimal]$imp
                    Sig = (Get-FirmaLoop $sist $mon $cue $fecha.Date $com $mov $signo $acc $imp)
                })
        }
        catch { $malas.Add("fila $n : $($_.Exception.Message)") }
    }
    Write-Malformadas $malas $csv.Archivo
    if ($rows.Count -eq 0) { throw "Loop sin filas válidas: $($csv.Archivo)" }
    $mn = $rows[0].Fecha; $mx = $rows[0].Fecha
    foreach ($x in $rows) { if ($x.Fecha -lt $mn) { $mn = $x.Fecha }; if ($x.Fecha -gt $mx) { $mx = $x.Fecha } }
    Write-Log ("{0}: {1} partidas válidas, fechas {2} a {3}." -f $csv.Archivo, $rows.Count, (Format-Fecha $mn), (Format-Fecha $mx))
    return [pscustomobject]@{ Rows = $rows; Min = $mn; Max = $mx; Archivo = $csv.Archivo }
}

function Read-Pendientes {
    param([string]$Path)
    $csv = Import-CsvTolerante -Path $Path
    $req = @('sistcod', 'moncod', 'cuecod', 'movcod', 'comnro', 'importedeb', 'fecorigen', 'fecproceso')
    Test-Columnas $csv.Header $req $csv.Archivo
    $fmts = @('d/M/yyyy HH:mm:ss.fff', 'd/M/yyyy H:mm:ss.fff', 'd/M/yyyy HH:mm:ss', 'd/M/yyyy H:mm:ss', 'd/M/yyyy')
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $malas = New-Object 'System.Collections.Generic.List[string]'
    $n = 1
    foreach ($r in $csv.Rows) {
        $n++
        try {
            $mon = Convert-Entero $r.moncod
            $cue = Convert-Entero $r.cuecod
            $imp = Convert-Decimal $r.importedeb
            $com = Format-Comnro $r.comnro
            $fecha = Convert-Fecha $r.fecorigen $fmts
            if ($null -eq $fecha) { $fecha = Convert-Fecha $r.fecproceso $fmts }
            if ($null -eq $mon -or $null -eq $imp -or $com -eq '') { throw 'moncod/importedeb/comnro inválido' }
            if ($null -eq $fecha) { throw "fecorigen/fecproceso inválida '$($r.fecorigen)'" }
            if ($null -eq $cue) { $cue = [long]0 }
            $rows.Add([pscustomobject]@{
                    Mon = [int]$mon; Cue = [long]$cue; Comnro = $com; Importe = [Math]::Abs([decimal]$imp); Fecha = $fecha.Date
                    Sist = (Convert-Entero $r.sistcod); Mov = (Convert-Entero $r.movcod)
                })
        }
        catch { $malas.Add("fila $n : $($_.Exception.Message)") }
    }
    Write-Malformadas $malas $csv.Archivo
    $noDeb = @($rows | Where-Object { $_.Sist -ne 5 -or $_.Mov -ne 25030 }).Count
    if ($noDeb -gt 0) { Write-Log "$($csv.Archivo): $noDeb pendiente(s) que no son sistcod 5 / movcod 25030." 'WARN' }
    Write-Log ("{0}: {1} débitos pendientes válidos." -f $csv.Archivo, $rows.Count)
    return [pscustomobject]@{ Rows = $rows; Archivo = $csv.Archivo }
}

function Read-Transferencias {
    param([string]$Path)
    $csv = Import-CsvTolerante -Path $Path
    $nuevo = ($csv.Header -contains 'Comprobante') -or ($csv.Header -contains 'Monto')
    $viejo = $csv.Header -contains 'nro_comprobante'
    if ($nuevo) {
        Test-Columnas $csv.Header @('Comprobante', 'Creada', 'Ejecutada', 'Estado', 'CBU origen', 'Monto', 'Moneda') $csv.Archivo
    }
    elseif ($viejo) {
        Test-Columnas $csv.Header @('nro_comprobante', 'holder_cbu', 'amount', 'status', 'fecha_ejecucion', 'execution_mode') $csv.Archivo
        Write-Log "$($csv.Archivo): formato viejo de transferencias (solo execution_mode = QNET)."
    }
    else { throw "Formato de transferencias no reconocido en $($csv.Archivo): faltan Comprobante/Monto o nro_comprobante." }

    $fmts = @('yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm:ss.fff', "yyyy-MM-dd'T'HH:mm:ss", 'yyyy-MM-dd')
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $malas = New-Object 'System.Collections.Generic.List[string]'
    $descartadas = 0
    $n = 1
    foreach ($r in $csv.Rows) {
        $n++
        try {
            if ($nuevo) {
                $com = Format-Comnro $r.Comprobante; $cbuRaw = $r.'CBU origen'; $monto = Convert-Decimal $r.Monto
                $estado = ([string]$r.Estado).Trim().ToUpper(); $ejec = Convert-Fecha $r.Ejecutada $fmts
                $creada = Convert-Fecha $r.Creada $fmts; $moneda = ([string]$r.Moneda).Trim().ToUpper()
            }
            else {
                if (([string]$r.execution_mode).Trim().ToUpper() -ne 'QNET') { $descartadas++; continue }
                $com = Format-Comnro $r.nro_comprobante; $cbuRaw = $r.holder_cbu; $monto = Convert-Decimal $r.amount
                $estado = ([string]$r.status).Trim().ToUpper(); $ejec = Convert-Fecha $r.fecha_ejecucion $fmts
                $creada = $null; $moneda = ''
            }
            if ($com -eq '') { throw 'comprobante vacío' }
            if ($null -eq $monto) { throw 'monto inválido' }
            if ($estado -ne 'APPROVED' -and $estado -ne 'FAILED') { throw "estado desconocido '$estado'" }
            $cbu = ([string]$cbuRaw) -replace '\D', ''
            $cuenta = $null; $tipoCta = $null
            if ($cbu.Length -ge 21) {
                $tipoCta = $cbu.Substring(8, 2)      # dígitos 9 y 10
                $cuenta = [long]$cbu.Substring(10, 11) # caracteres 11 a 21
            }
            if ($moneda -eq 'ARS') { $mon = 0 }
            elseif ($moneda -eq 'USD') { $mon = 2 }
            elseif ($moneda -eq '' -and $tipoCta -eq '02') { $mon = 0 }
            elseif ($moneda -eq '' -and $tipoCta -eq '06') { $mon = 2 }
            else { throw "moneda no reconocida '$moneda'" }
            $cbuOk = $null
            if ($null -ne $tipoCta) {
                if ($mon -eq 0) { $cbuOk = ($tipoCta -eq '02') } else { $cbuOk = ($tipoCta -eq '06') }
            }
            $fecha = $null
            if ($null -ne $ejec) { $fecha = $ejec.Date } elseif ($null -ne $creada) { $fecha = $creada.Date }
            $rows.Add([pscustomobject]@{
                    Comnro = $com; Estado = $estado; Ejecutada = $ejec; Fecha = $fecha; Monto = [decimal]$monto
                    Mon = [int]$mon; Cuenta = $cuenta; TipoCta = $tipoCta; CbuOk = $cbuOk
                })
        }
        catch { $malas.Add("fila $n : $($_.Exception.Message)") }
    }
    Write-Malformadas $malas $csv.Archivo
    if ($descartadas -gt 0) { Write-Log "$($csv.Archivo): $descartadas fila(s) con execution_mode distinto de QNET descartadas." }
    $corte = $null
    foreach ($x in $rows) { if ($null -ne $x.Ejecutada -and ($null -eq $corte -or $x.Ejecutada -gt $corte)) { $corte = $x.Ejecutada } }
    $ap = @($rows | Where-Object { $_.Estado -eq 'APPROVED' }).Count
    Write-Log ("{0}: {1} operaciones ({2} APPROVED, {3} FAILED). Corte: {4}." -f $csv.Archivo, $rows.Count, $ap, ($rows.Count - $ap), $(if ($corte) { $corte.ToString('dd/MM/yyyy HH:mm', $script:Inv) } else { 'sin datos' }))
    return [pscustomobject]@{ Rows = $rows; Corte = $corte; Archivo = $csv.Archivo }
}

# ---------------------------------------------------------------------------
# Combinación de Loop (base + período nuevo)
# ---------------------------------------------------------------------------
function Merge-Loop {
    param($Loop, $Base)
    $info = [pscustomobject]@{ Rows = $Loop.Rows; Min = $Loop.Min; Max = $Loop.Max; Faltantes = 0; Combinado = $false }
    if ($null -eq $Base) {
        Write-Log ("Loop sin base. Período: {0} a {1}." -f (Format-Fecha $Loop.Min), (Format-Fecha $Loop.Max))
        return $info
    }
    if ($Loop.Min -le $Base.Min) {
        Write-Log 'La fecha mínima de Loop no es mayor que la de la base: se usa solo Loop.' 'WARN'
        return $info
    }
    # Las filas de la base desde la fecha mínima de Loop deben estar en Loop (todas las columnas)
    $cuenta = @{}
    foreach ($x in $Loop.Rows) {
        if ($cuenta.ContainsKey($x.Sig)) { $cuenta[$x.Sig]++ } else { $cuenta[$x.Sig] = 1 }
    }
    $faltan = New-Object 'System.Collections.Generic.List[object]'
    foreach ($b in $Base.Rows) {
        if ($b.Fecha -lt $Loop.Min) { continue }
        if ($cuenta.ContainsKey($b.Sig) -and $cuenta[$b.Sig] -gt 0) { $cuenta[$b.Sig]-- }
        else { $faltan.Add($b) }
    }
    if ($faltan.Count -gt 0) {
        Write-Log "$($faltan.Count) fila(s) de la base con fecha >= $(Format-Fecha $Loop.Min) no están en Loop:" 'WARN'
        $faltan | Select-Object -First 50 | ForEach-Object {
            Write-Log ("  comprobante {0}, mon {1}, sist {2}, {3}, importe {4}" -f $_.Comnro, $_.Mon, $_.Sist, (Format-Fecha $_.Fecha), (Format-Num $_.Imp)) 'WARN'
        }
        if ($faltan.Count -gt 50) { Write-Log "  ... y $($faltan.Count - 50) más." 'WARN' }
    }
    else { Write-Log 'Control de base: todas las filas de la base desde la fecha mínima de Loop están en Loop.' 'OK' }
    $res = New-Object 'System.Collections.Generic.List[object]'
    $antes = 0
    foreach ($b in $Base.Rows) { if ($b.Fecha -lt $Loop.Min) { $res.Add($b); $antes++ } }
    foreach ($x in $Loop.Rows) { $res.Add($x) }
    $mn = $res[0].Fecha; $mx = $res[0].Fecha
    foreach ($x in $res) { if ($x.Fecha -lt $mn) { $mn = $x.Fecha }; if ($x.Fecha -gt $mx) { $mx = $x.Fecha } }
    Write-Log ("Loop combinado: {0} filas de la base (anteriores a {1}) + {2} de Loop. Período: {3} a {4}." -f $antes, (Format-Fecha $Loop.Min), $Loop.Rows.Count, (Format-Fecha $mn), (Format-Fecha $mx))
    return [pscustomobject]@{ Rows = $res; Min = $mn; Max = $mx; Faltantes = $faltan.Count; Combinado = $true }
}

# ---------------------------------------------------------------------------
# Conciliación
# ---------------------------------------------------------------------------
function New-Hallazgo {
    param([string]$Tipo, [string]$Comprobante, $Cuenta, [int]$MonCod, $Fecha, $Importe, [string]$Estado, $DebPend, [string]$Detalle, [bool]$EsExcepcion = $true)
    return [pscustomobject]@{
        Tipo = $Tipo; Comprobante = $Comprobante; Cuenta = $Cuenta; MonCod = $MonCod; Fecha = $Fecha
        Importe = $Importe; Estado = $Estado; DebPend = $DebPend; Detalle = $Detalle; EsExcepcion = $EsExcepcion
    }
}

function Get-Conciliacion {
    param($Loop, $Pendientes, $Transfers)
    if ($Loop.Rows.Count -eq 0) { throw 'Loop no tiene partidas.' }
    $loopMin = $Loop.Min.Date; $loopMax = $Loop.Max.Date
    $corte = $Transfers.Corte
    $hall = New-Object 'System.Collections.Generic.List[object]'

    # Índices de pendientes y transfers
    $pDict = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $pDup = 0
    foreach ($p in $Pendientes.Rows) {
        $k = "$($p.Mon)|$($p.Comnro)"
        if ($pDict.ContainsKey($k)) { $pDup++ } else { $pDict[$k] = $p }
    }
    if ($pDup -gt 0) { Write-Log "$pDup débito(s) pendiente(s) repetido(s) para la misma clave (se usa el primero)." 'WARN' }
    $tDict = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $tDup = 0
    foreach ($t in $Transfers.Rows) {
        if ($tDict.ContainsKey($t.Comnro)) {
            $tDup++
            if ($tDict[$t.Comnro].Estado -ne 'APPROVED' -and $t.Estado -eq 'APPROVED') { $tDict[$t.Comnro] = $t }
        }
        else { $tDict[$t.Comnro] = $t }
    }
    if ($tDup -gt 0) { Write-Log "$tDup operación(es) de transfers repetida(s) por comprobante (se prioriza APPROVED)." 'WARN' }

    # Agrupar Loop por moncod + comnro
    $grupos = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach ($r in $Loop.Rows) {
        $k = "$($r.Mon)|$($r.Comnro)"
        if (-not $grupos.ContainsKey($k)) { $grupos[$k] = New-Object 'System.Collections.Generic.List[object]' }
        $grupos[$k].Add($r)
    }

    $comps = New-Object 'System.Collections.Generic.List[object]'
    $cDict = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $porComnro = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $idxDeb = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach ($k in $grupos.Keys) {
        $g = $grupos[$k]
        $deb = [decimal]0; $cred = [decimal]0; $neto = [decimal]0; $fd = $null; $fc = $null; $cta = $null
        foreach ($p in $g) {
            $neto += $p.Imp
            if ($p.Signo -lt 0) { $deb += $p.Acc; if ($null -eq $fd -or $p.Fecha -lt $fd) { $fd = $p.Fecha } }
            elseif ($p.Signo -gt 0) { $cred += $p.Acc; if ($null -eq $fc -or $p.Fecha -lt $fc) { $fc = $p.Fecha } }
            if ($p.Sist -eq 5 -and $null -eq $cta) { $cta = $p.Cue }
        }
        $f = $fc; if ($null -eq $f) { $f = $fd }; if ($null -eq $f) { $f = $g[0].Fecha }
        $c = [pscustomobject]@{
            Key = $k; MonCod = $g[0].Mon; Moneda = (Get-NombreMoneda $g[0].Mon); Comnro = $g[0].Comnro; Partidas = $g.Count
            Debito = $deb; Credito = $cred; Neto = $neto; FechaDeb = $fd; FechaCred = $fc; Fecha = $f
            Cuenta = $cta; CuentaLoop = $cta; Marca = ''; Estado = 'Sin operación'; PendImporte = $null; Detalle = ''
        }
        $comps.Add($c)
        $cDict[$k] = $c
        if (-not $porComnro.ContainsKey($c.Comnro)) { $porComnro[$c.Comnro] = New-Object 'System.Collections.Generic.List[object]' }
        $porComnro[$c.Comnro].Add($c)
        if ($null -ne $cta -and $deb -gt 0) {
            $ik = '{0}|{1}|{2}' -f $c.MonCod, $cta, $deb.ToString('0.00########', $script:Inv)
            if (-not $idxDeb.ContainsKey($ik)) { $idxDeb[$ik] = New-Object 'System.Collections.Generic.List[string]' }
            $idxDeb[$ik].Add($c.Comnro)
        }
    }

    # Marcas
    foreach ($c in $comps) {
        $pend = $null
        if ($pDict.ContainsKey($c.Key)) { $pend = $pDict[$c.Key] }
        $ref = $c.Credito; if ($ref -eq 0) { $ref = $c.Debito }
        if ($c.Partidas -gt 2) {
            $c.Marca = 'Más de 2 partidas'
            $c.Detalle = '{0} partidas; neto {1}' -f $c.Partidas, (Format-Importe $c.Neto $c.MonCod)
        }
        elseif ($c.Partidas -eq 1 -and $null -ne $pend) {
            $c.PendImporte = $pend.Importe
            if ([Math]::Abs($pend.Importe - $c.Credito) -le $script:Tol) {
                $c.Marca = 'Débito pendiente'
                $c.Detalle = 'Débito pendiente de ' + (Format-Importe $pend.Importe $c.MonCod)
            }
            else {
                $c.Marca = 'Pendiente con importe distinto'
                $c.Detalle = 'Pendiente {0}; crédito {1}' -f (Format-Importe $pend.Importe $c.MonCod), (Format-Importe $c.Credito $c.MonCod)
                $ik = '{0}|{1}|{2}' -f $c.MonCod, $pend.Cue, $pend.Importe.ToString('0.00########', $script:Inv)
                if ($idxDeb.ContainsKey($ik)) {
                    $otros = @($idxDeb[$ik] | Where-Object { $_ -ne $c.Comnro })
                    if ($otros.Count -gt 0) {
                        $c.Detalle += "; mismo importe y cuenta que el débito del $($otros[0]) (posible débito duplicado o cruzado)"
                    }
                }
            }
        }
        elseif ($c.Partidas -eq 1) {
            $c.Marca = 'Una partida'
            $c.Detalle = '1 partida de {0} sin débito pendiente' -f (Format-Importe $ref $c.MonCod)
        }
        elseif ([Math]::Abs($c.Neto) -gt $script:Tol) {
            $c.Marca = 'Importes distintos'
            $c.Detalle = 'Débito {0}; crédito {1}; neto {2}' -f (Format-Importe $c.Debito $c.MonCod), (Format-Importe $c.Credito $c.MonCod), (Format-Importe $c.Neto $c.MonCod)
        }
        else { $c.Marca = 'OK' }
        if ($null -eq $c.Cuenta -and $null -ne $pend) { $c.Cuenta = $pend.Cue }
    }

    # Cruce de comprobantes de Loop con transfers
    foreach ($c in $comps) {
        $t = $null
        if ($tDict.ContainsKey($c.Comnro)) { $t = $tDict[$c.Comnro] }
        $ref = $c.Credito; if ($ref -eq 0) { $ref = $c.Debito }
        if ($null -ne $t) {
            $c.Estado = $t.Estado
            if ($null -eq $c.Cuenta) { $c.Cuenta = $t.Cuenta }
            if ($t.Estado -eq 'FAILED') {
                $hall.Add((New-Hallazgo 'FAILED contabilizado' $c.Comnro $c.Cuenta $c.MonCod $c.Fecha $ref 'FAILED' $c.PendImporte 'La operación figura FAILED en transfers y tiene asiento en Loop'))
            }
            if ([Math]::Abs($t.Monto - $ref) -gt $script:Tol) {
                $hall.Add((New-Hallazgo 'Importe distinto a transfer' $c.Comnro $c.Cuenta $c.MonCod $c.Fecha $ref $t.Estado $c.PendImporte ('Loop {0}; transfer {1}' -f (Format-Importe $ref $c.MonCod), (Format-Importe $t.Monto $t.Mon))))
            }
            if ($t.Mon -ne $c.MonCod) {
                $hall.Add((New-Hallazgo 'Moneda distinta a transfer' $c.Comnro $c.Cuenta $c.MonCod $c.Fecha $ref $t.Estado $c.PendImporte ('Loop {0}; transfer {1}' -f (Get-NombreMoneda $c.MonCod), (Get-NombreMoneda $t.Mon))))
            }
            if ($null -ne $c.CuentaLoop -and $null -ne $t.Cuenta -and $c.CuentaLoop -ne $t.Cuenta) {
                $hall.Add((New-Hallazgo 'Cuenta distinta a transfer' $c.Comnro $c.CuentaLoop $c.MonCod $c.Fecha $ref $t.Estado $c.PendImporte ('Loop {0}; transfer {1}' -f $c.CuentaLoop, $t.Cuenta)))
            }
        }
        elseif ($c.Partidas -eq 2 -and $null -ne $corte -and $c.Fecha -lt $corte.Date) {
            $hall.Add((New-Hallazgo 'Sin operación en transfers' $c.Comnro $c.Cuenta $c.MonCod $c.Fecha $ref 'Sin operación' $null ('No figura en transfers (corte {0})' -f (Format-Fecha $corte))))
        }
    }

    # Cruce de transfers contra Loop
    $det = New-Object 'System.Collections.Generic.List[object]'
    $enLoop = 0; $failSin = 0; $apSin = 0; $apPend = 0; $apFuera = 0; $cbuMal = 0
    foreach ($t in $Transfers.Rows) {
        if ($t.CbuOk -eq $false) {
            $cbuMal++
            $esp = '02'; if ($t.Mon -eq 2) { $esp = '06' }
            $hall.Add((New-Hallazgo 'CBU origen no corresponde a la moneda' $t.Comnro $t.Cuenta $t.Mon $t.Fecha $t.Monto $t.Estado $null ("Dígitos 9 y 10 del CBU origen = '{0}', se esperaba '{1}'" -f $t.TipoCta, $esp)))
        }
        if ($porComnro.ContainsKey($t.Comnro)) { $enLoop++; $det.Add([pscustomobject]@{ Fecha = $t.Fecha; Clase = 'EnLoop' }); continue }
        if ($t.Estado -eq 'FAILED') { $failSin++; $det.Add([pscustomobject]@{ Fecha = $t.Fecha; Clase = 'FailedSin' }); continue }
        if ($null -eq $t.Fecha) { $apSin++; $det.Add([pscustomobject]@{ Fecha = $null; Clase = 'ApprovedSin' }); $hall.Add((New-Hallazgo 'APPROVED sin registro' $t.Comnro $t.Cuenta $t.Mon $null $t.Monto 'APPROVED' $null 'APPROVED sin fecha y sin asiento en Loop')); continue }
        if ($t.Fecha -ge $loopMax) { $apPend++; $det.Add([pscustomobject]@{ Fecha = $t.Fecha; Clase = 'ApprovedPend' }) }
        elseif ($t.Fecha -lt $loopMin) { $apFuera++; $det.Add([pscustomobject]@{ Fecha = $t.Fecha; Clase = 'ApprovedFuera' }) }
        else {
            $apSin++
            $det.Add([pscustomobject]@{ Fecha = $t.Fecha; Clase = 'ApprovedSin' })
            $hall.Add((New-Hallazgo 'APPROVED sin registro' $t.Comnro $t.Cuenta $t.Mon $t.Fecha $t.Monto 'APPROVED' $null 'APPROVED dentro del período de Loop sin asiento'))
        }
    }

    # Validaciones de los débitos pendientes
    $pendFuturos = 0
    foreach ($k in $pDict.Keys) {
        $p = $pDict[$k]
        $c = $null
        if ($cDict.ContainsKey($k)) { $c = $cDict[$k] }
        $t = $null
        if ($tDict.ContainsKey($p.Comnro)) { $t = $tDict[$p.Comnro] }
        $estT = 'Sin operación'; if ($null -ne $t) { $estT = $t.Estado }
        if ($null -eq $c) {
            if ($p.Fecha -gt $loopMax) { $pendFuturos++ }
            else { $hall.Add((New-Hallazgo 'Pendiente sin asiento en Loop' $p.Comnro $p.Cue $p.Mon $p.Fecha $p.Importe $estT $p.Importe 'El débito pendiente no tiene comprobante en Loop')) }
        }
        elseif ($c.Partidas -ge 2) {
            $hall.Add((New-Hallazgo 'Pendiente sobre comprobante completo' $p.Comnro $p.Cue $p.Mon $c.Fecha $p.Importe $estT $p.Importe ('El comprobante ya tiene {0} partidas en Loop (posible débito duplicado)' -f $c.Partidas)))
        }
        if ($null -ne $t) {
            if ([Math]::Abs($t.Monto - $p.Importe) -gt $script:Tol) {
                $hall.Add((New-Hallazgo 'Pendiente: importe distinto a transfer' $p.Comnro $p.Cue $p.Mon $p.Fecha $p.Importe $estT $p.Importe ('Pendiente {0}; transfer {1}' -f (Format-Importe $p.Importe $p.Mon), (Format-Importe $t.Monto $t.Mon))))
            }
            if ($null -ne $t.Cuenta -and $p.Cue -ne $t.Cuenta) {
                $hall.Add((New-Hallazgo 'Pendiente: cuenta distinta a transfer' $p.Comnro $p.Cue $p.Mon $p.Fecha $p.Importe $estT $p.Importe ('Pendiente {0}; transfer {1}' -f $p.Cue, $t.Cuenta)))
            }
        }
    }

    # Totales por moneda
    $tot = @{}
    foreach ($m in @(0, 2)) { $tot[$m] = [pscustomobject]@{ Neto = [decimal]0; Pendientes = [decimal]0; SinCubrir = [decimal]0 } }
    foreach ($c in $comps) {
        if (-not $tot.ContainsKey($c.MonCod)) { $tot[$c.MonCod] = [pscustomobject]@{ Neto = [decimal]0; Pendientes = [decimal]0; SinCubrir = [decimal]0 } }
        $tot[$c.MonCod].Neto += $c.Neto
        if ($c.Marca -eq 'Débito pendiente') { $tot[$c.MonCod].Pendientes += $c.PendImporte }
    }
    foreach ($m in @($tot.Keys)) { $tot[$m].SinCubrir = $tot[$m].Neto - $tot[$m].Pendientes }

    $ordenados = @($comps | Sort-Object MonCod, Fecha, { $_.Comnro.PadLeft(20) })
    Write-Log "Conciliación: $($comps.Count) comprobantes de Loop, $($pDict.Count) pendientes, $($Transfers.Rows.Count) operaciones de transfers."
    return [pscustomobject]@{
        Comprobantes = $ordenados; Hallazgos = $hall; Totales = $tot; TransfersDet = $det
        LoopMin = $loopMin; LoopMax = $loopMax; Corte = $corte
        Cruce = [pscustomobject]@{
            Total = $Transfers.Rows.Count; EnLoop = $enLoop; FailedSinRegistro = $failSin; ApprovedSinRegistro = $apSin
            ApprovedPendientes = $apPend; ApprovedFuera = $apFuera; CbuMoneda = $cbuMal
        }
        PendientesPosteriores = $pendFuturos
        PendientesTotal = $pDict.Count
    }
}

function Get-Excepciones {
    # Une las marcas de cada comprobante con los hallazgos del cruce y ordena por tipo
    param($Conc, [switch]$Silencioso)
    $res = New-Object 'System.Collections.Generic.List[object]'
    $marcas = @('Más de 2 partidas', 'Una partida', 'Importes distintos', 'Pendiente con importe distinto', 'Débito pendiente')
    foreach ($c in $Conc.Comprobantes) {
        if ($marcas -notcontains $c.Marca) { continue }
        $ref = $c.Credito; if ($ref -eq 0) { $ref = $c.Debito }
        $esExc = ($c.Marca -ne 'Débito pendiente')
        $res.Add((New-Hallazgo $c.Marca $c.Comnro $c.Cuenta $c.MonCod $c.Fecha $ref $c.Estado $c.PendImporte $c.Detalle $esExc))
    }
    foreach ($h in $Conc.Hallazgos) { $res.Add($h) }
    $orden = $script:OrdenTipos
    $sorted = @($res | Sort-Object @{ Expression = { $i = [array]::IndexOf($orden, $_.Tipo); if ($i -lt 0) { 99 } else { $i } } }, @{ Expression = { $_.MonCod } }, @{ Expression = { if ($null -eq $_.Fecha) { [datetime]::MaxValue } else { $_.Fecha } } }, @{ Expression = { $_.Comprobante } })
    $n = @($sorted | Where-Object { $_.EsExcepcion }).Count
    if (-not $Silencioso) { Write-Log "Excepciones detectadas: $n (más $(@($sorted).Count - $n) comprobantes con débito pendiente listados)." }
    return , $sorted
}

function Get-PendientesPorDia {
    # Saldo de pendientes por fecha de crédito y moneda, con acumulado
    param($Conc)
    $pend = @($Conc.Comprobantes | Where-Object { $_.Marca -eq 'Débito pendiente' })
    $fechas = @($pend | ForEach-Object { $_.Fecha } | Sort-Object -Unique)
    $filas = New-Object 'System.Collections.Generic.List[object]'
    $acP = [decimal]0; $acU = [decimal]0
    foreach ($f in $fechas) {
        $p = @($pend | Where-Object { $_.Fecha -eq $f -and $_.MonCod -eq 0 })
        $u = @($pend | Where-Object { $_.Fecha -eq $f -and $_.MonCod -eq 2 })
        $sp = [decimal]0; foreach ($x in $p) { $sp += $x.PendImporte }
        $su = [decimal]0; foreach ($x in $u) { $su += $x.PendImporte }
        $acP += $sp; $acU += $su
        $filas.Add([pscustomobject]@{ Fecha = $f; CantPesos = $p.Count; Pesos = $sp; AcumPesos = $acP; CantUsd = $u.Count; Usd = $su; AcumUsd = $acU })
    }
    return , $filas
}

# ---------------------------------------------------------------------------
# Resumen (estructura común para Excel y CSV)
# ---------------------------------------------------------------------------
function Get-BloquesResumen {
    param($Conc, $Meta)
    $bloques = New-Object 'System.Collections.Generic.List[object]'
    $marcas = @('OK', 'Débito pendiente', 'Pendiente con importe distinto', 'Una partida', 'Importes distintos', 'Más de 2 partidas')
    $filas = @()
    foreach ($m in $marcas) { $filas += , @($m, @($Conc.Comprobantes | Where-Object { $_.Marca -eq $m }).Count) }
    $filas += , @('Total de comprobantes', $Conc.Comprobantes.Count)
    $bloques.Add(@{ Titulo = 'Comprobantes por marca'; Cols = @('Marca', 'Cantidad'); Filas = $filas; Formula = $false })
    $x = $Conc.Cruce
    $bloques.Add(@{ Titulo = 'Transferencias Loop BankService'; Cols = @('Concepto', 'Cantidad'); Formula = $false; Filas = @(
                @('Total de operaciones', $x.Total),
                @('Encontradas en Loop', $x.EnLoop),
                @('FAILED sin registro', $x.FailedSinRegistro),
                @('APPROVED sin registro', $x.ApprovedSinRegistro),
                @('APPROVED pendientes de contabilizar', $x.ApprovedPendientes),
                @('APPROVED fuera del período', $x.ApprovedFuera),
                @('Pendientes posteriores al último día de Loop', $Conc.PendientesPosteriores)
            )
    })
    $tp = $Conc.Totales[0]; $tu = $Conc.Totales[2]
    $bloques.Add(@{ Titulo = 'Desbalance'; Cols = @('Concepto', 'Pesos', 'USD'); Formula = $true; Filas = @(
                @('Neto en Loop', $tp.Neto, $tu.Neto),
                @('Débitos pendientes', $tp.Pendientes, $tu.Pendientes),
                @('Sin cubrir', $tp.SinCubrir, $tu.SinCubrir)
            )
    })
    $bloques.Add(@{ Titulo = 'Archivos y períodos'; Cols = @('Concepto', 'Valor'); Formula = $false; Filas = @(
                @('Loop', $Meta.LoopArchivo),
                @('Loop base', $Meta.BaseArchivo),
                @('Filas de la base ausentes en Loop', $Meta.Faltantes),
                @('Débitos pendientes', $Meta.PendArchivo),
                @('Transferencias', $Meta.TransArchivo),
                @('Período de Loop', (Format-Periodo $Conc.LoopMin $Conc.LoopMax)),
                @('Corte de transfers', $(if ($Conc.Corte) { $Conc.Corte.ToString('dd/MM/yyyy HH:mm', $script:Inv) } else { '' }))
            )
    })
    return , $bloques
}

# ---------------------------------------------------------------------------
# Excel (ImportExcel) con respaldo en CSV
# ---------------------------------------------------------------------------
function Export-ExcelReport {
    param($Conc, $Excepciones, $PendDia, $Meta, [string]$Base)
    $xlsx = $Base + '.xlsx'
    $bloques = Get-BloquesResumen $Conc $Meta

    $filasExc = @($Excepciones | ForEach-Object {
            [pscustomobject][ordered]@{
                'Tipo' = $_.Tipo; 'Comprobante' = $_.Comprobante; 'Cuenta' = $_.Cuenta; 'Moneda' = (Get-NombreMoneda $_.MonCod)
                'Fecha' = $_.Fecha; 'Importe' = $_.Importe; 'Estado transfer' = $_.Estado; 'Débito pendiente' = $_.DebPend; 'Detalle' = $_.Detalle
            } })
    $filasComp = @($Conc.Comprobantes | ForEach-Object {
            [pscustomobject][ordered]@{
                'Clave' = ('{0}-{1}' -f $_.MonCod, $_.Comnro); 'Moneda' = $_.Moneda; 'Comprobante' = $_.Comnro; 'Partidas' = $_.Partidas
                'Débito' = $_.Debito; 'Crédito' = $_.Credito; 'Neto' = $_.Neto; 'Fecha débito' = $_.FechaDeb; 'Fecha crédito' = $_.FechaCred
                'Cuenta' = $_.Cuenta; 'Marca' = $_.Marca; 'Estado transfer' = $_.Estado
            } })

    if ($null -eq $script:TieneIE) { $script:TieneIE = [bool](Get-Module -ListAvailable -Name ImportExcel) }
    if (-not $script:TieneIE) {
        if (-not $script:AvisoIE) {
            Write-Log 'No está el módulo ImportExcel. Instalalo con: Install-Module ImportExcel -Scope CurrentUser' 'WARN'
            Write-Log 'Como respaldo se exportan las hojas como CSV.' 'WARN'
            $script:AvisoIE = $true
        }
        $resumen = foreach ($b in $bloques) {
            foreach ($f in $b.Filas) {
                $o = [ordered]@{ 'Bloque' = $b.Titulo }
                for ($i = 0; $i -lt $b.Cols.Count; $i++) { $o[$b.Cols[$i]] = $f[$i] }
                [pscustomobject]$o
            }
        }
        $fmtFecha = { param($v) if ($null -eq $v) { '' } else { ([datetime]$v).ToString('yyyy-MM-dd', $script:Inv) } }
        $csvs = @{
            'Resumen'           = $resumen
            'Excepciones'       = @($filasExc | ForEach-Object { $_.Fecha = (& $fmtFecha $_.Fecha); $_ })
            'Pendientes por dia' = @($PendDia | ForEach-Object {
                    [pscustomobject][ordered]@{ 'Fecha del crédito' = (& $fmtFecha $_.Fecha); 'Comprobantes en pesos' = $_.CantPesos; 'Pesos' = $_.Pesos; 'Acumulado en pesos' = $_.AcumPesos
                        'Comprobantes en USD' = $_.CantUsd; 'USD' = $_.Usd; 'Acumulado en USD' = $_.AcumUsd } })
            'Transferencias_Loop_IBS'     = @($filasComp | ForEach-Object { $_.'Fecha débito' = (& $fmtFecha $_.'Fecha débito'); $_.'Fecha crédito' = (& $fmtFecha $_.'Fecha crédito'); $_ })
        }
        $out = @()
        foreach ($nombre in $csvs.Keys) {
            $ruta = '{0}_{1}.csv' -f $Base, ($nombre -replace ' ', '_')
            if (@($csvs[$nombre]).Count -eq 0) { Set-Content -LiteralPath $ruta -Value '' -Encoding UTF8 }
            else { @($csvs[$nombre]) | Export-Csv -LiteralPath $ruta -NoTypeInformation -Encoding UTF8 }
            $out += $ruta
            Write-Log "CSV generado: $ruta"
        }
        return $out
    }

    Import-Module ImportExcel
    # Si ya existe el Excel del mismo día (corrida anterior) se reemplaza; si está abierto en Excel, avisa
    if (Test-Path -LiteralPath $xlsx) {
        try { Remove-Item -LiteralPath $xlsx -Force }
        catch { throw "No se pudo reemplazar $xlsx. Cerralo si lo tenés abierto en Excel y volvé a ejecutar." }
    }
    $pkg = New-Object OfficeOpenXml.ExcelPackage (New-Object System.IO.FileInfo $xlsx)
    try {
        $encabezadosExc = 'Tipo', 'Comprobante', 'Cuenta', 'Moneda', 'Fecha', 'Importe', 'Estado transfer', 'Débito pendiente', 'Detalle'

        # Excepciones
        if ($filasExc.Count -gt 0) {
            $null = Export-Excel -ExcelPackage $pkg -WorksheetName 'Excepciones' -InputObject $filasExc -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
        }
        else {
            $w = $pkg.Workbook.Worksheets.Add('Excepciones')
            for ($i = 0; $i -lt $encabezadosExc.Count; $i++) { $w.Cells[1, ($i + 1)].Value = $encabezadosExc[$i]; $w.Cells[1, ($i + 1)].Style.Font.Bold = $true }
            $w.View.FreezePanes(2, 1)
        }
        $we = $pkg.Workbook.Worksheets['Excepciones']
        $ultima = [Math]::Max(2, $filasExc.Count + 1)
        $we.Cells[2, 5, $ultima, 5].Style.Numberformat.Format = 'dd/mm/yyyy'
        $we.Cells[2, 6, $ultima, 6].Style.Numberformat.Format = '#,##0.00'
        $we.Cells[2, 8, $ultima, 8].Style.Numberformat.Format = '#,##0.00'
        $we.Cells[2, 2, $ultima, 3].Style.Numberformat.Format = '0'
        try { $we.Cells[$we.Dimension.Address].AutoFitColumns() } catch { }

        # Pendientes por día (fórmulas sobre Excepciones)
        $wp = $pkg.Workbook.Worksheets.Add('Pendientes por día')
        $enc = 'Fecha del crédito', 'Comprobantes en pesos', 'Pesos', 'Acumulado en pesos', 'Comprobantes en USD', 'USD', 'Acumulado en USD'
        for ($i = 0; $i -lt $enc.Count; $i++) { $wp.Cells[1, ($i + 1)].Value = $enc[$i]; $wp.Cells[1, ($i + 1)].Style.Font.Bold = $true }
        $r = 2
        foreach ($d in $PendDia) {
            $wp.Cells[$r, 1].Value = $d.Fecha
            $wp.Cells[$r, 1].Style.Numberformat.Format = 'dd/mm/yyyy'
            $criterio = "Excepciones!`$A`$2:`$A`$$ultima,""Débito pendiente"",Excepciones!`$E`$2:`$E`$$ultima,`$A$r,Excepciones!`$D`$2:`$D`$$ultima,"
            $wp.Cells[$r, 2].Formula = "COUNTIFS(${criterio}""Pesos"")"
            $wp.Cells[$r, 3].Formula = "SUMIFS(Excepciones!`$F`$2:`$F`$$ultima,$criterio""Pesos"")"
            $wp.Cells[$r, 5].Formula = "COUNTIFS(${criterio}""USD"")"
            $wp.Cells[$r, 6].Formula = "SUMIFS(Excepciones!`$F`$2:`$F`$$ultima,$criterio""USD"")"
            if ($r -eq 2) { $wp.Cells[$r, 4].Formula = "C$r"; $wp.Cells[$r, 7].Formula = "F$r" }
            else { $wp.Cells[$r, 4].Formula = "D$($r - 1)+C$r"; $wp.Cells[$r, 7].Formula = "G$($r - 1)+F$r" }
            $r++
        }
        $wp.Cells[$r, 1].Value = 'Total'
        if ($r -gt 2) {
            foreach ($col in 'B', 'C', 'E', 'F') { $wp.Cells["$col$r"].Formula = "SUM(${col}2:${col}$($r - 1))" }
        }
        else { foreach ($col in 'B', 'C', 'E', 'F') { $wp.Cells["$col$r"].Value = 0 } }
        $wp.Cells[$r, 1, $r, 7].Style.Font.Bold = $true
        foreach ($col in 3, 4, 6, 7) { $wp.Cells[2, $col, $r, $col].Style.Numberformat.Format = '#,##0.00' }
        $wp.View.FreezePanes(2, 1)
        try { $wp.Cells[$wp.Dimension.Address].AutoFitColumns() } catch { }

        # Comprobantes
        if ($filasComp.Count -gt 0) {
            $null = Export-Excel -ExcelPackage $pkg -WorksheetName 'Transferencias Loop IBS' -InputObject $filasComp -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
            $wc = $pkg.Workbook.Worksheets['Transferencias Loop IBS']
            $uc = $filasComp.Count + 1
            $wc.Cells[2, 5, $uc, 7].Style.Numberformat.Format = '#,##0.00'
            $wc.Cells[2, 8, $uc, 9].Style.Numberformat.Format = 'dd/mm/yyyy'
            $wc.Cells[2, 3, $uc, 3].Style.Numberformat.Format = '0'
            $wc.Cells[2, 10, $uc, 10].Style.Numberformat.Format = '0'
            try { $wc.Cells[$wc.Dimension.Address].AutoFitColumns() } catch { }
        }
        else { $null = $pkg.Workbook.Worksheets.Add('Transferencias Loop IBS') }

        # Resumen
        $ws = $pkg.Workbook.Worksheets.Add('Resumen')
        $ws.Cells[1, 1].Value = 'Conciliación Loop - IBS - BankService' + $(if ($Meta.Dia) { ' - ' + (Format-Fecha $Meta.Dia) } else { '' })
        $ws.Cells[1, 1].Style.Font.Bold = $true; $ws.Cells[1, 1].Style.Font.Size = 14
        $ws.Cells[2, 1].Value = 'Generado: ' + (Get-Date).ToString('dd/MM/yyyy HH:mm', $script:Inv)
        $r = 4
        foreach ($b in $bloques) {
            $ws.Cells[$r, 1].Value = $b.Titulo; $ws.Cells[$r, 1].Style.Font.Bold = $true; $ws.Cells[$r, 1].Style.Font.Size = 12
            $r++
            for ($i = 0; $i -lt $b.Cols.Count; $i++) { $ws.Cells[$r, ($i + 1)].Value = $b.Cols[$i]; $ws.Cells[$r, ($i + 1)].Style.Font.Bold = $true }
            $r++
            $primera = $r
            foreach ($f in $b.Filas) {
                for ($i = 0; $i -lt $f.Count; $i++) {
                    $celda = $ws.Cells[$r, ($i + 1)]
                    if ($null -ne $f[$i]) { $celda.Value = $f[$i] }
                    if ($f[$i] -is [decimal]) { $celda.Style.Numberformat.Format = '#,##0.00' }
                }
                $r++
            }
            if ($b.Formula) {
                # Sin cubrir = neto - pendientes (fórmula)
                $u = $r - 1
                $ws.Cells[$u, 2].Formula = "B$($primera)-B$($primera + 1)"
                $ws.Cells[$u, 3].Formula = "C$($primera)-C$($primera + 1)"
                $ws.Cells[$u, 1, $u, 3].Style.Font.Bold = $true
            }
            $r++
        }
        try { $ws.Cells[$ws.Dimension.Address].AutoFitColumns() } catch { }
        $pkg.Workbook.Worksheets.MoveToStart('Resumen')
        try { $pkg.Workbook.Calculate() } catch { Write-Log 'No se pudieron precalcular las fórmulas; Excel las calculará al abrir.' 'WARN' }
        $pkg.Save()
    }
    finally { $pkg.Dispose() }
    Write-Log "Excel generado: $xlsx" 'OK'
    return @($xlsx)
}

# ---------------------------------------------------------------------------
# Dashboard HTML
# ---------------------------------------------------------------------------
function Get-CssBase {
    return @'
<style>
:root{--bg:#f4f6f8;--card:#fff;--fg:#1b2430;--muted:#5d6978;--line:#dde2e8;--alert:#c62828;--alertbg:#fdecea;
--s0:#2a9d8f;--s1:#3d6fd1;--s2:#e08a1e;--s3:#c62828;--head:#eef1f5}
@media (prefers-color-scheme:dark){:root{--bg:#12161c;--card:#1b222b;--fg:#e6ebf1;--muted:#97a3b2;--line:#2c3643;--alert:#ff8a80;--alertbg:#3a1f1f;
--s0:#4fc3b4;--s1:#7ea4f2;--s2:#f0aa4d;--s3:#ff7b72;--head:#232c37}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.45 -apple-system,"Segoe UI",Roboto,Arial,sans-serif}
main{max-width:1200px;margin:0 auto;padding:20px 16px 48px}
h1{margin:0 0 4px;font-size:22px}h2{margin:28px 0 10px;font-size:16px}
.muted{color:var(--muted)}.meta{display:flex;flex-wrap:wrap;gap:6px 24px;color:var(--muted)}
.resumen{margin:12px 0;padding:12px 14px;background:var(--card);border:1px solid var(--line);border-radius:8px;font-weight:600}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px}
.tile{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:10px 12px}
.tile .lbl{color:var(--muted);font-size:12px}.tile .val{font-size:20px;font-weight:700;margin-top:2px;word-break:break-word}
.tile .sub{color:var(--muted);font-size:12px}
.tile.alert{border-color:var(--alert);background:var(--alertbg)}.tile.alert .val{color:var(--alert)}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px}
.scroll{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums}
th,td{padding:6px 10px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
td.n,th.n{text-align:right}td.det{white-space:normal;min-width:240px}
thead th{background:var(--head);position:sticky;top:0}
th.sort{cursor:pointer;user-select:none}th.sort:after{content:" \2195";color:var(--muted)}
th.asc:after{content:" \25B2"}th.desc:after{content:" \25BC"}
tfoot td{font-weight:700}
.btns{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
button.f{background:var(--card);color:var(--fg);border:1px solid var(--line);border-radius:16px;padding:4px 12px;cursor:pointer;font:inherit}
button.f.on{background:var(--s1);border-color:var(--s1);color:#fff}
input[type=search]{width:100%;max-width:360px;padding:7px 10px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--fg);font:inherit;margin-bottom:10px}
svg{width:100%;height:auto;display:block}svg .gl{stroke:var(--line);stroke-width:1}svg .ax{fill:var(--muted);font-size:11px}
.s0{fill:var(--s0)}.s1{fill:var(--s1)}.s2{fill:var(--s2)}.s3{fill:var(--s3)}
.legend{display:flex;flex-wrap:wrap;gap:6px 16px;margin-bottom:6px}.sw{display:inline-block;width:11px;height:11px;border-radius:2px;margin-right:5px;vertical-align:-1px}
i.s0{background:var(--s0)}i.s1{background:var(--s1)}i.s2{background:var(--s2)}i.s3{background:var(--s3)}
#tip{position:fixed;display:none;pointer-events:none;background:#111;color:#fff;padding:4px 8px;border-radius:4px;font-size:12px;z-index:9}
@media (max-width:600px){h1{font-size:18px}.tile .val{font-size:17px}}
label{display:block;font-size:12px;color:var(--muted)}input[type=date]{padding:6px 8px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--fg);font:inherit}
.filtros{display:flex;flex-wrap:wrap;gap:10px;align-items:end}a{color:var(--s1)}caption{padding:8px;color:var(--muted);text-align:left}
</style></head><body><main>
'@
}

function New-SvgBarras {
    param($Conc)
    $dias = @{}
    foreach ($c in $Conc.Comprobantes) {
        $idx = -1
        if ($c.Marca -eq 'OK') { $idx = 0 }
        elseif ($c.Marca -eq 'Débito pendiente' -or $c.Marca -eq 'Pendiente con importe distinto') { $idx = 1 }
        elseif ($c.Marca -eq 'Una partida') { $idx = 2 }
        elseif ($c.Marca -eq 'Importes distintos') { $idx = 3 }
        if ($idx -lt 0) { continue }
        $k = $c.Fecha.ToString('yyyyMMdd', $script:Inv)
        if (-not $dias.ContainsKey($k)) { $dias[$k] = [pscustomobject]@{ Fecha = $c.Fecha; V = @(0, 0, 0, 0) } }
        $dias[$k].V[$idx]++
    }
    $keys = @($dias.Keys | Sort-Object)
    if ($keys.Count -eq 0) { return '<p class="muted">Sin datos para graficar.</p>' }
    $nombres = 'Correctos', 'Débito pendiente', 'Una partida', 'Importes distintos'
    $W = 900; $H = 320; $ml = 46; $mr = 10; $mt = 14; $mb = 44
    $pw = $W - $ml - $mr; $ph = $H - $mt - $mb
    $max = 1
    foreach ($k in $keys) { $s = 0; foreach ($v in $dias[$k].V) { $s += $v }; if ($s -gt $max) { $max = $s } }
    $f = { param($x) [string]::Format($script:Inv, '{0:0.##}', [double]$x) }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox=""0 0 $W $H"" role=""img"" aria-label=""Comprobantes por fecha y estado"" preserveAspectRatio=""xMidYMid meet"">")
    foreach ($frac in 0, 0.5, 1) {
        $y = $mt + $ph - ($frac * $ph)
        $val = [Math]::Round($frac * $max)
        [void]$sb.Append("<line class=""gl"" x1=""$ml"" x2=""$($W - $mr)"" y1=""$(& $f $y)"" y2=""$(& $f $y)""/>")
        [void]$sb.Append("<text class=""ax"" x=""$($ml - 6)"" y=""$(& $f ($y + 4))"" text-anchor=""end"">$val</text>")
    }
    $slot = $pw / $keys.Count
    $bw = [Math]::Min(40, $slot * 0.7)
    $paso = [Math]::Ceiling($keys.Count / 18)
    $i = 0
    foreach ($k in $keys) {
        $d = $dias[$k]
        $x = $ml + $slot * $i + ($slot - $bw) / 2
        $base = $mt + $ph
        for ($s = 0; $s -lt 4; $s++) {
            $v = $d.V[$s]
            if ($v -le 0) { continue }
            $h = $v / $max * $ph
            $base -= $h
            $tip = Get-Enc ('{0} - {1}: {2}' -f (Format-Fecha $d.Fecha), $nombres[$s], $v)
            [void]$sb.Append("<rect class=""s$s"" x=""$(& $f $x)"" y=""$(& $f $base)"" width=""$(& $f $bw)"" height=""$(& $f $h)"" data-tip=""$tip""/>")
        }
        if ($i % $paso -eq 0) {
            [void]$sb.Append("<text class=""ax"" x=""$(& $f ($x + $bw / 2))"" y=""$($H - 24)"" text-anchor=""middle"">$($d.Fecha.ToString('dd/MM', $script:Inv))</text>")
        }
        $i++
    }
    [void]$sb.Append('</svg>')
    $leyenda = New-Object System.Text.StringBuilder
    [void]$leyenda.Append('<div class="legend">')
    for ($s = 0; $s -lt 4; $s++) { [void]$leyenda.Append("<span><i class=""sw s$s""></i>$(Get-Enc $nombres[$s])</span>") }
    [void]$leyenda.Append('</div>')
    return $leyenda.ToString() + $sb.ToString()
}

function Export-HtmlDashboard {
    param($Conc, $Excepciones, $PendDia, $Meta, [string]$Base, [switch]$EmbedirDatos)
    $ruta = $Base + '.html'
    $sb = New-Object System.Text.StringBuilder
    $ap = { param($t) [void]$sb.AppendLine($t) }

    $cnt = { param($m) @($Conc.Comprobantes | Where-Object { $_.Marca -eq $m }).Count }
    $nOk = & $cnt 'OK'; $nPend = & $cnt 'Débito pendiente'; $nUna = & $cnt 'Una partida'; $nDist = & $cnt 'Importes distintos'
    $nDos = @($Conc.Comprobantes | Where-Object { $_.Partidas -eq 2 }).Count
    $excs = @($Excepciones | Where-Object { $_.EsExcepcion })
    $tp = $Conc.Totales[0]; $tu = $Conc.Totales[2]
    $x = $Conc.Cruce
    $periodo = Format-Periodo $Conc.LoopMin $Conc.LoopMax
    $corteTxt = 'sin datos'; if ($Conc.Corte) { $corteTxt = $Conc.Corte.ToString('dd/MM/yyyy HH:mm', $script:Inv) }
    $frase = '{0} comprobantes con débito pendiente, {1} excepciones y {2} / {3} sin cubrir.' -f $nPend, $excs.Count, (Format-Importe $tp.SinCubrir 0), (Format-Importe $tu.SinCubrir 2)

    $tile = {
        param([string]$Rotulo, $Valor, [bool]$Alerta, [string]$Sub)
        $cls = 'tile'; if ($Alerta) { $cls = 'tile alert' }
        $s = "<div class=""$cls""><div class=""lbl"">$(Get-Enc $Rotulo)</div><div class=""val"">$(Get-Enc ([string]$Valor))</div>"
        if ($Sub) { $s += "<div class=""sub"">$(Get-Enc $Sub)</div>" }
        return $s + '</div>'
    }
    $dinero = { param($m) [Math]::Abs($m) -gt $script:Tol }

    & $ap '<!DOCTYPE html>'
    & $ap '<html lang="es"><head><meta charset="utf-8">'
    & $ap '<meta name="viewport" content="width=device-width, initial-scale=1">'
    & $ap '<title>Conciliación Loop</title>'
    & $ap (Get-CssBase)
    $tituloH1 = 'Conciliación Loop - IBS - BankService'
    if ($Meta.Dia) { $tituloH1 += ' - ' + (Format-Fecha $Meta.Dia) }
    & $ap ('<h1>{0}</h1>' -f (Get-Enc $tituloH1))
    & $ap ('<div class="meta"><span>Período de Loop: <b>{0}</b></span><span>Corte de transfers: <b>{1}</b></span><span>Generado: {2}</span></div>' -f (Get-Enc $periodo), (Get-Enc $corteTxt), (Get-Enc (Get-Date).ToString('dd/MM/yyyy HH:mm', $script:Inv)))
    & $ap ('<div class="resumen">{0}</div>' -f (Get-Enc $frase))

    & $ap '<h2>Transferencias Loop IBS</h2><div class="grid">'
    & $ap (& $tile 'Con 2 partidas' $nDos $false ('Correctos: {0}' -f $nOk))
    & $ap (& $tile 'Con débito pendiente' $nPend $false '')
    & $ap (& $tile 'Con una partida' $nUna ($nUna -gt 0) '')
    & $ap (& $tile 'Con importes distintos' $nDist ($nDist -gt 0) '')
    & $ap (& $tile 'Excepciones' $excs.Count ($excs.Count -gt 0) '')
    & $ap '</div>'

    & $ap '<h2>Desbalance por moneda</h2><div class="grid">'
    & $ap (& $tile 'Neto en Loop - Pesos' (Format-Importe $tp.Neto 0) $false '')
    & $ap (& $tile 'Pendientes - Pesos' (Format-Importe $tp.Pendientes 0) $false '')
    & $ap (& $tile 'Sin cubrir - Pesos' (Format-Importe $tp.SinCubrir 0) (& $dinero $tp.SinCubrir) '')
    & $ap (& $tile 'Neto en Loop - USD' (Format-Importe $tu.Neto 2) $false '')
    & $ap (& $tile 'Pendientes - USD' (Format-Importe $tu.Pendientes 2) $false '')
    & $ap (& $tile 'Sin cubrir - USD' (Format-Importe $tu.SinCubrir 2) (& $dinero $tu.SinCubrir) '')
    & $ap '</div>'

    & $ap '<h2>Transferencias Loop BankService</h2><div class="grid">'
    & $ap (& $tile 'Operaciones' $x.Total $false '')
    & $ap (& $tile 'Encontradas en Loop' $x.EnLoop $false '')
    & $ap (& $tile 'FAILED sin registro' $x.FailedSinRegistro $false '')
    & $ap (& $tile 'APPROVED sin registro' $x.ApprovedSinRegistro ($x.ApprovedSinRegistro -gt 0) '')
    & $ap (& $tile 'APPROVED pendientes de contabilizar' $x.ApprovedPendientes $false 'No son excepción')
    & $ap (& $tile 'APPROVED fuera del período' $x.ApprovedFuera $false '')
    & $ap '</div>'

    & $ap '<h2>Transferencias Loop por fecha</h2><div class="card">'
    & $ap (New-SvgBarras $Conc)
    & $ap '</div>'

    & $ap '<h2>Saldo de débitos pendientes por día</h2><div class="card scroll">'
    & $ap '<table><thead><tr><th>Fecha del crédito</th><th class="n">Comprobantes $</th><th class="n">Pesos</th><th class="n">Acumulado $</th><th class="n">Comprobantes US$</th><th class="n">USD</th><th class="n">Acumulado US$</th></tr></thead><tbody>'
    $sumCp = 0; $sumCu = 0; $sumP = [decimal]0; $sumU = [decimal]0
    foreach ($d in $PendDia) {
        $sumCp += $d.CantPesos; $sumCu += $d.CantUsd; $sumP += $d.Pesos; $sumU += $d.Usd
        $imp = { param($v, $m, $c) if ($c -eq 0) { '' } else { Get-Enc (Format-Importe $v $m) } }
        & $ap ('<tr><td>{0}</td><td class="n">{1}</td><td class="n">{2}</td><td class="n">{3}</td><td class="n">{4}</td><td class="n">{5}</td><td class="n">{6}</td></tr>' -f
            (Get-Enc (Format-Fecha $d.Fecha)), $d.CantPesos, (& $imp $d.Pesos 0 $d.CantPesos), (Get-Enc (Format-Importe $d.AcumPesos 0)), $d.CantUsd, (& $imp $d.Usd 2 $d.CantUsd), (Get-Enc (Format-Importe $d.AcumUsd 2)))
    }
    & $ap ('</tbody><tfoot><tr><td>Total</td><td class="n">{0}</td><td class="n">{1}</td><td class="n"></td><td class="n">{2}</td><td class="n">{3}</td><td class="n"></td></tr></tfoot></table></div>' -f $sumCp, (Get-Enc (Format-Importe $sumP 0)), $sumCu, (Get-Enc (Format-Importe $sumU 2)))

    & $ap '<h2>Excepciones a revisar</h2><div class="card">'
    if ($excs.Count -eq 0) { & $ap '<p>Sin excepciones.</p>' }
    else {
        $tipos = @($excs | ForEach-Object { $_.Tipo } | Select-Object -Unique)
        & $ap ('<div class="btns" id="filtros"><button class="f on" data-tipo="">Todos ({0})</button>' -f $excs.Count)
        foreach ($t in $tipos) {
            $n = @($excs | Where-Object { $_.Tipo -eq $t }).Count
            & $ap ('<button class="f" data-tipo="{0}">{1} ({2})</button>' -f (Get-Enc $t), (Get-Enc $t), $n)
        }
        & $ap '</div><input type="search" id="q" placeholder="Buscar por comprobante, cuenta o fecha">'
        & $ap '<div class="scroll"><table id="exc"><thead><tr><th class="sort">Tipo</th><th class="sort">Comprobante</th><th class="sort">Cuenta</th><th class="sort">Moneda</th><th class="sort">Fecha</th><th class="sort n">Importe</th><th class="sort">Estado transfer</th><th>Detalle</th></tr></thead><tbody>'
        foreach ($e in $excs) {
            $fv = ''; if ($null -ne $e.Fecha) { $fv = ([datetime]$e.Fecha).ToString('yyyyMMdd', $script:Inv) }
            $iv = ''; if ($null -ne $e.Importe) { $iv = ([decimal]$e.Importe).ToString('0.00', $script:Inv) }
            $it = ''; if ($null -ne $e.Importe) { $it = Format-Importe $e.Importe $e.MonCod }
            $q = ('{0} {1} {2} {3}' -f $e.Comprobante, $e.Cuenta, (Format-Fecha $e.Fecha), $fv).ToLower()
            & $ap ('<tr data-tipo="{0}" data-q="{1}"><td>{2}</td><td data-v="{3}">{3}</td><td data-v="{4}">{4}</td><td>{5}</td><td data-v="{6}">{7}</td><td class="n" data-v="{8}">{9}</td><td>{10}</td><td class="det">{11}</td></tr>' -f
                (Get-Enc $e.Tipo), (Get-Enc $q), (Get-Enc $e.Tipo), (Get-Enc ([string]$e.Comprobante)), (Get-Enc ([string]$e.Cuenta)), (Get-Enc (Get-NombreMoneda $e.MonCod)),
                $fv, (Get-Enc (Format-Fecha $e.Fecha)), $iv, (Get-Enc $it), (Get-Enc $e.Estado), (Get-Enc $e.Detalle))
        }
        & $ap '</tbody></table></div>'
    }
    & $ap '</div>'
    & $ap ('<p class="muted">Archivos: {0}, {1}, {2}{3}. Los pendientes se informan por fecha del crédito en Loop.</p>' -f (Get-Enc $Meta.LoopArchivo), (Get-Enc $Meta.PendArchivo), (Get-Enc $Meta.TransArchivo), $(if ($Meta.BaseArchivo) { ', base ' + (Get-Enc $Meta.BaseArchivo) } else { '' }))
    & $ap '<div id="tip"></div>'
    if ($EmbedirDatos) { & $ap ('<script type="application/json" id="datos-conciliacion">' + (Get-DatosDia $Conc $Excepciones $PendDia $Meta.Dia) + '</script>') }
    & $ap @'
<script>
(function(){
var tip=document.getElementById('tip');
document.querySelectorAll('[data-tip]').forEach(function(el){
  function show(e){tip.textContent=el.getAttribute('data-tip');tip.style.display='block';tip.style.left=Math.min(e.clientX+12,window.innerWidth-200)+'px';tip.style.top=(e.clientY+14)+'px';}
  el.addEventListener('mousemove',show);el.addEventListener('click',show);
  el.addEventListener('mouseleave',function(){tip.style.display='none';});
});
var tabla=document.getElementById('exc');if(!tabla)return;
var filas=Array.prototype.slice.call(tabla.tBodies[0].rows);
var tipo='',q='';
function aplicar(){filas.forEach(function(r){
  var ok=(!tipo||r.getAttribute('data-tipo')===tipo)&&(!q||r.getAttribute('data-q').indexOf(q)>=0);r.hidden=!ok;});}
document.querySelectorAll('#filtros button').forEach(function(b){b.addEventListener('click',function(){
  document.querySelectorAll('#filtros button').forEach(function(x){x.classList.remove('on');});
  b.classList.add('on');tipo=b.getAttribute('data-tipo');aplicar();});});
document.getElementById('q').addEventListener('input',function(e){q=e.target.value.trim().toLowerCase();aplicar();});
function valor(r,i){var c=r.cells[i];var v=c.hasAttribute('data-v')?c.getAttribute('data-v'):c.textContent;return v;}
Array.prototype.forEach.call(tabla.tHead.rows[0].cells,function(th,i){
  if(!th.classList.contains('sort'))return;
  th.addEventListener('click',function(){
    var asc=!th.classList.contains('asc');
    Array.prototype.forEach.call(th.parentNode.cells,function(x){x.classList.remove('asc','desc');});
    th.classList.add(asc?'asc':'desc');
    filas.sort(function(a,b){
      var x=valor(a,i),y=valor(b,i),nx=parseFloat(x),ny=parseFloat(y),r;
      if(x!==''&&y!==''&&!isNaN(nx)&&!isNaN(ny)&&isFinite(x)&&isFinite(y))r=nx-ny;else r=x.localeCompare(y,'es');
      return asc?r:-r;});
    filas.forEach(function(r){tabla.tBodies[0].appendChild(r);});
  });});
})();
</script></main></body></html>
'@
    [System.IO.File]::WriteAllText($ruta, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    Write-Log "Dashboard generado: $ruta" 'OK'
    return $ruta
}

function Get-DatosDia {
    # Datos del día en JSON (sin importes sensibles) para que el dashboard general los pueda leer
    param($Conc, $Excepciones, $PendDia, [datetime]$Dia)
    $n = { param($m) @($Conc.Comprobantes | Where-Object { $_.Marca -eq $m }).Count }
    $d2 = { param($v) ([decimal]$v).ToString('0.00', $script:Inv) }
    $tp = $Conc.Totales[0]; $tu = $Conc.Totales[2]
    $x = $Conc.Cruce
    $pd = $null
    if ($PendDia.Count -gt 0) { $pd = $PendDia[0] }
    $cantP = 0; $cantU = 0; $sumP = [decimal]0; $sumU = [decimal]0
    if ($null -ne $pd) { $cantP = $pd.CantPesos; $cantU = $pd.CantUsd; $sumP = $pd.Pesos; $sumU = $pd.Usd }
    [object[]]$exc = @($Excepciones | Where-Object { $_.EsExcepcion } | ForEach-Object {
            $imp = ''
            if ($null -ne $_.Importe) { $imp = & $d2 $_.Importe }
            @{ t = $_.Tipo; c = [string]$_.Comprobante; cta = [string]$_.Cuenta; m = [int]$_.MonCod; imp = $imp; est = [string]$_.Estado; det = [string]$_.Detalle }
        })
    $corte = ''
    if ($Conc.Corte) { $corte = $Conc.Corte.ToString('dd/MM/yyyy HH:mm', $script:Inv) }
    $o = [ordered]@{
        fecha   = $Dia.ToString('yyyy-MM-dd', $script:Inv)
        corte   = $corte
        marcas  = @{ ok = (& $n 'OK'); pend = (& $n 'Débito pendiente'); pendDist = (& $n 'Pendiente con importe distinto'); una = (& $n 'Una partida'); dist = (& $n 'Importes distintos'); mas2 = (& $n 'Más de 2 partidas') }
        dos     = @($Conc.Comprobantes | Where-Object { $_.Partidas -eq 2 }).Count
        pesos   = @{ neto = (& $d2 $tp.Neto); pend = (& $d2 $tp.Pendientes) }
        usd     = @{ neto = (& $d2 $tu.Neto); pend = (& $d2 $tu.Pendientes) }
        cruce   = @{ total = $x.Total; enLoop = $x.EnLoop; failedSin = $x.FailedSinRegistro; approvedSin = $x.ApprovedSinRegistro; approvedPend = $x.ApprovedPendientes; approvedFuera = $x.ApprovedFuera }
        pendDia = @{ cantP = $cantP; pesos = (& $d2 $sumP); cantU = $cantU; usd = (& $d2 $sumU) }
        exc     = $exc
    }
    $json = $o | ConvertTo-Json -Depth 6 -Compress
    return $json.Replace('<', '<')
}

function Get-ConciliacionDia {
    # Recorta la conciliación a un solo día (misma forma que Get-Conciliacion)
    param($Conc, [datetime]$Dia)
    $comps = @($Conc.Comprobantes | Where-Object { $_.Fecha -eq $Dia })
    $hall = New-Object 'System.Collections.Generic.List[object]'
    foreach ($h in $Conc.Hallazgos) {
        if ($null -ne $h.Fecha -and ([datetime]$h.Fecha).Date -eq $Dia) { $hall.Add($h) }
    }
    $det = @($Conc.TransfersDet | Where-Object { $_.Fecha -eq $Dia })
    $cuenta = { param($clase) @($det | Where-Object { $_.Clase -eq $clase }).Count }
    $tot = @{}
    foreach ($m in @(0, 2)) { $tot[$m] = [pscustomobject]@{ Neto = [decimal]0; Pendientes = [decimal]0; SinCubrir = [decimal]0 } }
    foreach ($c in $comps) {
        if (-not $tot.ContainsKey($c.MonCod)) { $tot[$c.MonCod] = [pscustomobject]@{ Neto = [decimal]0; Pendientes = [decimal]0; SinCubrir = [decimal]0 } }
        $tot[$c.MonCod].Neto += $c.Neto
        if ($c.Marca -eq 'Débito pendiente') { $tot[$c.MonCod].Pendientes += $c.PendImporte }
    }
    foreach ($m in @($tot.Keys)) { $tot[$m].SinCubrir = $tot[$m].Neto - $tot[$m].Pendientes }
    return [pscustomobject]@{
        Comprobantes = $comps; Hallazgos = $hall; Totales = $tot; LoopMin = $Dia; LoopMax = $Dia; Corte = $Conc.Corte
        Cruce = [pscustomobject]@{
            Total = $det.Count; EnLoop = (& $cuenta 'EnLoop'); FailedSinRegistro = (& $cuenta 'FailedSin'); ApprovedSinRegistro = (& $cuenta 'ApprovedSin')
            ApprovedPendientes = (& $cuenta 'ApprovedPend'); ApprovedFuera = (& $cuenta 'ApprovedFuera')
            CbuMoneda = @($hall | Where-Object { $_.Tipo -eq 'CBU origen no corresponde a la moneda' }).Count
        }
        PendientesPosteriores = 0; PendientesTotal = @($comps | Where-Object { $null -ne $_.PendImporte }).Count
    }
}

function Export-DashboardGeneral {
    # Dashboard general con filtro por fechas: lee los datos embebidos en los HTML diarios de <OutputDir>\Diario
    param([string]$OutputDir)
    $dirDia = Join-Path $OutputDir 'Diario'
    if (-not (Test-Path -LiteralPath $dirDia -PathType Container)) { throw "No existe la carpeta de reportes diarios: $dirDia" }
    $archivos = @(Get-ChildItem -LiteralPath $dirDia -Filter 'Conciliacion_Loop_*.html' | Where-Object { $_.Name -match '^Conciliacion_Loop_\d{8}\.html$' } | Sort-Object Name)
    if ($archivos.Count -eq 0) { throw "No hay HTML diarios (Conciliacion_Loop_yyyyMMdd.html) en $dirDia" }
    $items = New-Object 'System.Collections.Generic.List[string]'
    foreach ($a in $archivos) {
        try {
            $html = Get-Content -LiteralPath $a.FullName -Raw -Encoding UTF8
            $m = [regex]::Match($html, '<script type="application/json" id="datos-conciliacion">(.*?)</script>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            if (-not $m.Success) { throw 'no tiene el bloque de datos' }
            $json = $m.Groups[1].Value
            $null = $json | ConvertFrom-Json
            $xl = (Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($a.FullName, '.xlsx'))).ToString().ToLower()
            $items.Add('{"archivo":"Diario/' + $a.Name + '","xlsx":' + $xl + ',"datos":' + $json + '}')
        }
        catch { Write-Log ('No se pudo leer {0}: {1}' -f $a.Name, $_.Exception.Message) 'WARN' }
    }
    if ($items.Count -eq 0) { throw 'Ningún HTML diario tiene datos válidos.' }
    $ruta = Join-Path $OutputDir 'Conciliacion_Loop_General.html'
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="es"><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$sb.AppendLine('<title>Conciliación Loop - General</title>')
    [void]$sb.AppendLine((Get-CssBase))
    [void]$sb.AppendLine('<h1>Conciliación Loop - IBS - BankService (general)</h1>')
    [void]$sb.AppendLine(('<div class="meta"><span>Días cargados: <b>{0}</b></span><span>Generado: {1}</span></div>' -f $items.Count, (Get-Enc (Get-Date).ToString('dd/MM/yyyy HH:mm', $script:Inv))))
    [void]$sb.AppendLine(@'
<div class="card filtros" style="margin-top:12px">
<div><label for="desde">Desde</label><input type="date" id="desde"></div>
<div><label for="hasta">Hasta</label><input type="date" id="hasta"></div>
<button class="f" id="bTodo">Todo</button><button class="f" id="bUlt">Último día</button><button class="f" id="b7">Últimos 7 días</button>
</div>
<div class="resumen" id="frase"></div>
<h2>Transferencias Loop IBS</h2><div class="grid" id="t1"></div>
<h2>Desbalance por moneda</h2><div class="grid" id="t2"></div>
<h2>Transferencias Loop BankService</h2><div class="grid" id="t3"></div>
<h2>Transferencias Loop por fecha</h2><div class="card" id="chart"></div>
<h2>Saldo de débitos pendientes por día</h2><div class="card scroll"><table id="saldo"></table></div>
<h2>Detalle por día</h2><div class="card scroll"><table id="dias"></table></div>
<h2>Excepciones a revisar</h2><div class="card">
<div class="btns" id="filtros"></div>
<input type="search" id="q" placeholder="Buscar por comprobante, cuenta o fecha">
<div class="scroll"><table id="exc"></table></div></div>
<div id="tip"></div>
'@)
    [void]$sb.AppendLine('<script type="application/json" id="datos">[' + ($items -join ',') + ']</script>')
    [void]$sb.AppendLine(@'
<script>
(function(){
var D=JSON.parse(document.getElementById('datos').textContent).filter(function(x){return x&&x.datos&&/^\d{4}-\d{2}-\d{2}$/.test(x.datos.fecha);});
D.sort(function(a,b){return a.datos.fecha<b.datos.fecha?-1:(a.datos.fecha>b.datos.fecha?1:0);});
var el=function(id){return document.getElementById(id);};
var num=function(s){var n=parseFloat(s);return isNaN(n)?0:n;};
function fmt(n){var neg=n<0;n=Math.abs(n);var p=(Math.round(n*100)/100).toFixed(2).split('.');p[0]=p[0].replace(/\B(?=(\d{3})+(?!\d))/g,'.');return (neg&&(p[0]!=='0'||p[1]!=='00')?'-':'')+p[0]+','+p[1];}
function money(n,usd){return (usd?'US$ ':'$ ')+fmt(n);}
function fdate(i){var p=i.split('-');return p[2]+'/'+p[1]+'/'+p[0];}
function mk(tag,txt,cls){var e=document.createElement(tag);if(txt!==undefined&&txt!==null)e.textContent=txt;if(cls)e.className=cls;return e;}
var sel=[],tipo='',q='',sortCol=-1,sortAsc=true;
function tiles(id,arr){var c=el(id);c.innerHTML='';arr.forEach(function(t){var d=mk('div',null,'tile'+(t[2]?' alert':''));d.appendChild(mk('div',t[0],'lbl'));d.appendChild(mk('div',String(t[1]),'val'));if(t[3])d.appendChild(mk('div',t[3],'sub'));c.appendChild(d);});}
function bindTips(){var tip=el('tip');document.querySelectorAll('[data-tip]').forEach(function(e){
  function show(ev){tip.textContent=e.getAttribute('data-tip');tip.style.display='block';tip.style.left=Math.min(ev.clientX+12,window.innerWidth-200)+'px';tip.style.top=(ev.clientY+14)+'px';}
  e.addEventListener('mousemove',show);e.addEventListener('click',show);e.addEventListener('mouseleave',function(){tip.style.display='none';});});}
function chart(){
  var box=el('chart');
  if(!sel.length){box.textContent='Sin datos en el rango elegido.';return;}
  var names=['Correctos','Débito pendiente','Una partida','Importes distintos'];
  var W=900,H=320,ml=46,mr=10,mt=14,mb=44,pw=W-ml-mr,ph=H-mt-mb,max=1;
  var vals=sel.map(function(x){var m=x.datos.marcas;return [m.ok||0,(m.pend||0)+(m.pendDist||0),m.una||0,m.dist||0];});
  vals.forEach(function(v){var t=v[0]+v[1]+v[2]+v[3];if(t>max)max=t;});
  var slot=pw/sel.length,bw=Math.min(40,slot*0.7),paso=Math.ceil(sel.length/18),h='';
  [0,0.5,1].forEach(function(fr){var y=mt+ph-fr*ph;h+='<line class="gl" x1="'+ml+'" x2="'+(W-mr)+'" y1="'+y+'" y2="'+y+'"/><text class="ax" x="'+(ml-6)+'" y="'+(y+4)+'" text-anchor="end">'+Math.round(fr*max)+'</text>';});
  sel.forEach(function(x,i){var xx=ml+slot*i+(slot-bw)/2,base=mt+ph,v=vals[i],f=fdate(x.datos.fecha);
    for(var k=0;k<4;k++){if(v[k]<=0)continue;var hh=v[k]/max*ph;base-=hh;h+='<rect class="s'+k+'" x="'+xx+'" y="'+base+'" width="'+bw+'" height="'+hh+'" data-tip="'+f+' - '+names[k]+': '+v[k]+'"/>';}
    if(i%paso===0)h+='<text class="ax" x="'+(xx+bw/2)+'" y="'+(H-24)+'" text-anchor="middle">'+f.slice(0,5)+'</text>';});
  var leg='<div class="legend">';for(var k=0;k<4;k++)leg+='<span><i class="sw s'+k+'"></i>'+names[k]+'</span>';
  box.innerHTML=leg+'</div><svg viewBox="0 0 '+W+' '+H+'" role="img" aria-label="Transferencias Loop por fecha y estado">'+h+'</svg>';
  bindTips();
}
function tabla(t,heads,rows,foot){
  t.innerHTML='';var th=mk('thead'),tr=mk('tr');
  heads.forEach(function(h){tr.appendChild(mk('th',h[0],h[1]||''));});th.appendChild(tr);t.appendChild(th);
  var tb=mk('tbody');rows.forEach(function(r){var row=mk('tr');r.forEach(function(c,i){
    var td;if(c&&c.href){td=mk('td');var a=mk('a',c.t);a.href=c.href;td.appendChild(a);}else td=mk('td',c,(heads[i][1]||'').indexOf('n')>=0?'n':'');row.appendChild(td);});tb.appendChild(row);});t.appendChild(tb);
  if(foot){var tf=mk('tfoot'),fr=mk('tr');foot.forEach(function(c,i){fr.appendChild(mk('td',c,(heads[i][1]||'').indexOf('n')>=0?'n':''));});tf.appendChild(fr);t.appendChild(tf);}
}
function excepciones(){
  var rows=[];sel.forEach(function(x){(x.datos.exc||[]).forEach(function(e){rows.push({f:x.datos.fecha,t:e.t,c:e.c,cta:e.cta,m:e.m,imp:num(e.imp),est:e.est,det:e.det});});});
  var tipos={};rows.forEach(function(r){tipos[r.t]=(tipos[r.t]||0)+1;});
  if(tipo&&!tipos[tipo])tipo='';
  var fb=el('filtros');fb.innerHTML='';
  function boton(txt,val){var b=mk('button',txt,'f'+(tipo===val?' on':''));b.addEventListener('click',function(){tipo=val;excepciones();});fb.appendChild(b);}
  boton('Todos ('+rows.length+')','');Object.keys(tipos).forEach(function(k){boton(k+' ('+tipos[k]+')',k);});
  var vis=rows.filter(function(r){
    if(tipo&&r.t!==tipo)return false;
    if(q){var s=(r.c+' '+r.cta+' '+fdate(r.f)+' '+r.f).toLowerCase();if(s.indexOf(q)<0)return false;}return true;});
  var keys=['f','t','c','cta','m','imp','est'];
  var ord=sortCol>=0?keys[sortCol]:'f';
  vis.sort(function(a,b){var x=a[ord],y=b[ord],r;
    if(typeof x==='number'&&typeof y==='number')r=x-y;else r=String(x).localeCompare(String(y),'es');
    if(r===0)r=String(a.c).localeCompare(String(b.c),'es');return sortAsc?r:-r;});
  var heads=[['Día','sort'],['Tipo','sort'],['Comprobante','sort'],['Cuenta','sort'],['Moneda','sort'],['Importe','sort n'],['Estado transfer','sort'],['Detalle','']];
  var data=vis.map(function(r){return [fdate(r.f),r.t,r.c,r.cta,r.m===2?'USD':(r.m===0?'Pesos':'Moneda '+r.m),money(r.imp,r.m===2),r.est,r.det];});
  var t=el('exc');
  if(!data.length){t.innerHTML='';t.appendChild(mk('caption','Sin excepciones en el rango elegido.'));return;}
  tabla(t,heads,data);
  Array.prototype.forEach.call(t.tHead.rows[0].cells,function(c,i){
    if(!c.classList.contains('sort'))return;
    c.classList.add(sortCol===i?(sortAsc?'asc':'desc'):'x');
    c.addEventListener('click',function(){if(sortCol===i)sortAsc=!sortAsc;else{sortCol=i;sortAsc=true;}excepciones();});});
}
function render(){
  var d=el('desde').value,h=el('hasta').value;
  sel=D.filter(function(x){var f=x.datos.fecha;return (!d||f>=d)&&(!h||f<=h);});
  var s={ok:0,pend:0,pendDist:0,una:0,dist:0,mas2:0,dos:0,nExc:0,pNeto:0,pPend:0,uNeto:0,uPend:0,total:0,enLoop:0,failedSin:0,approvedSin:0,approvedPend:0,approvedFuera:0};
  sel.forEach(function(x){var o=x.datos,m=o.marcas,c=o.cruce;
    s.ok+=m.ok||0;s.pend+=m.pend||0;s.pendDist+=m.pendDist||0;s.una+=m.una||0;s.dist+=m.dist||0;s.mas2+=m.mas2||0;s.dos+=o.dos||0;s.nExc+=(o.exc||[]).length;
    s.pNeto+=num(o.pesos.neto);s.pPend+=num(o.pesos.pend);s.uNeto+=num(o.usd.neto);s.uPend+=num(o.usd.pend);
    s.total+=c.total||0;s.enLoop+=c.enLoop||0;s.failedSin+=c.failedSin||0;s.approvedSin+=c.approvedSin||0;s.approvedPend+=c.approvedPend||0;s.approvedFuera+=c.approvedFuera||0;});
  var sinP=s.pNeto-s.pPend,sinU=s.uNeto-s.uPend;
  var rango=sel.length?(sel.length===1?fdate(sel[0].datos.fecha):fdate(sel[0].datos.fecha)+' al '+fdate(sel[sel.length-1].datos.fecha)):'sin días';
  el('frase').textContent=sel.length+' día(s) ('+rango+'): '+s.pend+' comprobantes con débito pendiente, '+s.nExc+' excepciones y '+money(sinP)+' / '+money(sinU,true)+' sin cubrir.';
  tiles('t1',[['Con 2 partidas',s.dos,false,'Correctos: '+s.ok],['Con débito pendiente',s.pend,false],['Con una partida',s.una,s.una>0],['Con importes distintos',s.dist,s.dist>0],['Excepciones',s.nExc,s.nExc>0]]);
  tiles('t2',[['Neto en Loop - Pesos',money(s.pNeto),false],['Pendientes - Pesos',money(s.pPend),false],['Sin cubrir - Pesos',money(sinP),Math.abs(sinP)>0.005],
    ['Neto en Loop - USD',money(s.uNeto,true),false],['Pendientes - USD',money(s.uPend,true),false],['Sin cubrir - USD',money(sinU,true),Math.abs(sinU)>0.005]]);
  tiles('t3',[['Operaciones',s.total,false],['Encontradas en Loop',s.enLoop,false],['FAILED sin registro',s.failedSin,false],['APPROVED sin registro',s.approvedSin,s.approvedSin>0],
    ['APPROVED pendientes de contabilizar',s.approvedPend,false,'No son excepción'],['APPROVED fuera del período',s.approvedFuera,false]]);
  chart();
  var aP=0,aU=0,cP=0,cU=0,tP=0,tU=0;
  var filas=sel.map(function(x){var p=x.datos.pendDia;aP+=num(p.pesos);aU+=num(p.usd);cP+=p.cantP||0;cU+=p.cantU||0;
    return [fdate(x.datos.fecha),p.cantP||0,p.cantP?money(num(p.pesos)):'',money(aP),p.cantU||0,p.cantU?money(num(p.usd),true):'',money(aU,true)];});
  tabla(el('saldo'),[['Fecha del crédito'],['Comprobantes $','n'],['Pesos','n'],['Acumulado $','n'],['Comprobantes US$','n'],['USD','n'],['Acumulado US$','n']],filas,
    ['Total',cP,money(aP),'',cU,money(aU,true),'']);
  tabla(el('dias'),[['Día'],['Correctos','n'],['Débito pendiente','n'],['Una partida','n'],['Importes distintos','n'],['Excepciones','n'],['Sin cubrir $','n'],['Sin cubrir US$','n'],['HTML'],['Excel']],
    sel.map(function(x){var o=x.datos,m=o.marcas;return [fdate(o.fecha),m.ok||0,(m.pend||0)+(m.pendDist||0),m.una||0,m.dist||0,(o.exc||[]).length,money(num(o.pesos.neto)-num(o.pesos.pend)),money(num(o.usd.neto)-num(o.usd.pend),true),
      {t:'Abrir',href:x.archivo},x.xlsx?{t:'Abrir',href:x.archivo.replace(/\.html$/,'.xlsx')}:'-'];}));
  excepciones();
}
function fijar(d,h){el('desde').value=d;el('hasta').value=h;render();}
el('desde').min=el('hasta').min=D[0].datos.fecha;el('desde').max=el('hasta').max=D[D.length-1].datos.fecha;
el('desde').addEventListener('change',render);el('hasta').addEventListener('change',render);
el('bTodo').addEventListener('click',function(){fijar('','');});
el('bUlt').addEventListener('click',function(){var f=D[D.length-1].datos.fecha;fijar(f,f);});
el('b7').addEventListener('click',function(){fijar(D[Math.max(0,D.length-7)].datos.fecha,D[D.length-1].datos.fecha);});
el('q').addEventListener('input',function(e){q=e.target.value.trim().toLowerCase();excepciones();});
render();
})();
</script></main></body></html>
'@)
    [System.IO.File]::WriteAllText($ruta, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    Write-Log ('Dashboard general generado: {0} ({1} días)' -f $ruta, $items.Count) 'OK'
    return $ruta
}

# ---------------------------------------------------------------------------
# Descarga desde Sybase (paso 1 del modo todo en uno)
# ---------------------------------------------------------------------------
function Format-CampoCsv {
    # Valor en texto invariante, sin saltos de línea y con comillas si hace falta
    param($Valor)
    if ($null -eq $Valor -or $Valor -is [System.DBNull]) { return '' }
    if ($Valor -is [datetime]) { $t = $Valor.ToString('d/M/yyyy HH:mm:ss.fff', $script:Inv) }
    elseif ($Valor -is [decimal] -or $Valor -is [double] -or $Valor -is [single]) { $t = ([System.IFormattable]$Valor).ToString($null, $script:Inv) }
    elseif ($Valor -is [byte[]]) { $t = '' }
    else { $t = [string]$Valor }
    $t = ($t -replace '[\r\n]+', ' ').Trim()
    if ($t.IndexOfAny([char[]]@(',', '"')) -ge 0) { $t = '"' + $t.Replace('"', '""') + '"' }
    return $t
}

function Read-CredencialPopup {
    # Ventana emergente para pedir usuario y clave (la clave se muestra oculta)
    param([string]$Titulo = 'Conexión a Sybase', [string]$Usuario)
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    }
    catch { return Get-Credential -Message $Titulo }
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Titulo
    $f.StartPosition = 'CenterScreen'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false
    $f.TopMost = $true
    $f.ClientSize = New-Object System.Drawing.Size(330, 140)
    $lu = New-Object System.Windows.Forms.Label
    $lu.Text = 'Usuario:'; $lu.Location = New-Object System.Drawing.Point(15, 20); $lu.AutoSize = $true
    $tu = New-Object System.Windows.Forms.TextBox
    $tu.Location = New-Object System.Drawing.Point(90, 17); $tu.Size = New-Object System.Drawing.Size(220, 22)
    if ($Usuario) { $tu.Text = $Usuario; $tu.ReadOnly = $true; $tu.TabStop = $false }
    $lp = New-Object System.Windows.Forms.Label
    $lp.Text = 'Clave:'; $lp.Location = New-Object System.Drawing.Point(15, 58); $lp.AutoSize = $true
    $tp = New-Object System.Windows.Forms.TextBox
    $tp.Location = New-Object System.Drawing.Point(90, 55); $tp.Size = New-Object System.Drawing.Size(220, 22)
    $tp.UseSystemPasswordChar = $true
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Aceptar'; $ok.Location = New-Object System.Drawing.Point(150, 98); $ok.Size = New-Object System.Drawing.Size(75, 28)
    $ok.Add_Click({
            if ($tu.Text.Trim() -eq '') { [void][System.Windows.Forms.MessageBox]::Show('Ingresá el usuario.', $Titulo) }
            else { $f.DialogResult = [System.Windows.Forms.DialogResult]::OK; $f.Close() }
        }.GetNewClosure())
    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancelar'; $ca.Location = New-Object System.Drawing.Point(235, 98); $ca.Size = New-Object System.Drawing.Size(75, 28)
    $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.AcceptButton = $ok
    $f.CancelButton = $ca
    $f.Controls.AddRange(@($lu, $tu, $lp, $tp, $ok, $ca))
    $f.Add_Shown({ $f.Activate(); if ($tu.ReadOnly) { $tp.Focus() } else { $tu.Focus() } }.GetNewClosure())
    $res = $f.ShowDialog()
    if ($res -ne [System.Windows.Forms.DialogResult]::OK) { $f.Dispose(); throw 'Conexión cancelada por el usuario.' }
    $sec = New-Object System.Security.SecureString
    foreach ($ch in $tp.Text.ToCharArray()) { $sec.AppendChar($ch) }
    $sec.MakeReadOnly()
    $usuario = $tu.Text.Trim()
    $tp.Text = ''
    $f.Dispose()
    return (New-Object System.Management.Automation.PSCredential($usuario, $sec))
}

function Export-ConsultaCsv {
    # Ejecuta una consulta (solo lectura) y la vuelca a un CSV; devuelve la cantidad de filas
    param($Conexion, [string]$Sql, [string]$Ruta, [int]$TimeoutSeg)
    $dir = Split-Path -Parent $Ruta
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
    $cmd = $Conexion.CreateCommand()
    $cmd.CommandText = $Sql
    $cmd.CommandTimeout = $TimeoutSeg
    $reader = $cmd.ExecuteReader()
    $writer = New-Object System.IO.StreamWriter($Ruta, $false, (New-Object System.Text.UTF8Encoding($false)))
    $filas = 0
    try {
        $n = $reader.FieldCount
        if ($n -eq 0) { throw 'La consulta no devolvió columnas.' }
        $cols = for ($i = 0; $i -lt $n; $i++) { $reader.GetName($i) }
        $writer.Write(($cols -join ',') + "`r`n")
        while ($reader.Read()) {
            $campos = for ($i = 0; $i -lt $n; $i++) { Format-CampoCsv $reader.GetValue($i) }
            $writer.Write(($campos -join ',') + "`r`n")
            $filas++
        }
    }
    finally {
        $writer.Dispose()
        $reader.Dispose()
    }
    return $filas
}

function Get-MotivoNoHabil {
    # Devuelve el motivo si la fecha es fin de semana o feriado; $null si es día hábil
    param([datetime]$Dia)
    if ($Dia.DayOfWeek -eq [System.DayOfWeek]::Saturday -or $Dia.DayOfWeek -eq [System.DayOfWeek]::Sunday) { return 'fin de semana' }
    $k = $Dia.ToString('yyyy-MM-dd', $script:Inv)
    if ($script:Feriados.ContainsKey($k)) { return 'feriado: ' + $script:Feriados[$k] }
    return $null
}

function Invoke-Descarga {
    # Baja de Sybase (solo lectura) todos los débitos pendientes y los movimientos de Loop del período a CSV.
    # Fechas: desde el último día hábil anterior hasta ayer (lunes: desde el viernes). No se ejecuta en fines de semana ni feriados.
    param([string]$Carpeta, $FechaDia, [string]$Dsn, [string]$ConnectionString, [string]$Server, [int]$Port, [string]$Driver, [string]$Usuario, $Credencial, [int]$TimeoutSeg)
    if ($null -ne $FechaDia) {
        $desde = ([datetime]$FechaDia).Date
        $hasta = $desde.AddDays(1)
    }
    else {
        $hasta = (Get-Date).Date
        if ($hasta.Year -ne 2026) { Write-Log "La lista de feriados del script es de 2026; actualizala para $($hasta.Year)." 'WARN' }
        $motivo = Get-MotivoNoHabil $hasta
        if ($motivo) {
            Write-Log "Hoy no es día hábil ($motivo): no se ejecuta (use -Fecha para forzar un día)." 'WARN'
            return $null
        }
        # Desde el último día hábil anterior: cubre fines de semana y feriados intermedios (lunes = viernes, @ayer = @hoy - 3)
        $desde = $hasta.AddDays(-1)
        while (Get-MotivoNoHabil $desde) { $desde = $desde.AddDays(-1) }
    }
    $ultimo = $hasta.AddDays(-1)
    if ($desde -eq $ultimo) { $diaTxt = Format-Fecha $desde }
    else { $diaTxt = 'del {0} al {1}' -f (Format-Fecha $desde), (Format-Fecha $ultimo) }
    $fechasSql = "declare @hoy datetime, @ayer datetime`r`nselect @hoy  = '{0}'`r`nselect @ayer = '{1}'`r`n" -f $hasta.ToString('yyyyMMdd', $script:Inv), $desde.ToString('yyyyMMdd', $script:Inv)

    # Débitos pendientes: se consultan todos (sin filtro de fecha)
    $sqlPendientes = @"
select b.fecreal, b.tercod, b.succod, b.sistcod, b.moncod, b.cuecod, movcod, descdet, b.fecproceso, b.comnro, b.nromen, m.mentexto, b.importedeb, fecenlace, fecorigen
From Banksys..tgl_deb_pendientes b
inner join Banksys..tgt_transaccion a on b.paiscod = a.paiscod and
b.bcocod = a.bcocod and
b.sistcod = a.sistcod and
b.trancod = a.trancod
left outer join Banksys..tdoc_mensages m on b.nromen = m.nromen
Where
b.paiscod = 54 and
b.bcocod = 312 and
b.succod = 1 and
a.tipofunc in (0,1)
and a.movcod in (25030, 25031)

order by b.fecorigen, b.fecreal, b.fecenlace, b.cuecod
"@

    $sqlLoop = $fechasSql + @"

select  a.sistcod, a.moncod, a.cuecod, convert(varchar, a.fecmov, 103) 'Fecha_Mov', a.comnro, a.movcod, b.signo, b.accimp, b.signo * b.accimp 'Importe'
from tgl_movimientos a, tgl_movimientos_aper b
where
a.fecmov >= @ayer
and a.fecmov <  @hoy
and a.concod = ' '
and a.fecreal = b.fecreal
and a.tercod = b.tercod
and b.funcod = 1
and a.movcod in (25030, 25031)
order by a.comnro
"@

    if (-not $ConnectionString) {
        if ($Dsn) { $cs = "DSN=$Dsn;" }
        elseif ($Server) { $cs = "Driver={$Driver};Server=$Server;Port=$Port;" }
        else { throw 'Para bajar los archivos indicá -Dsn, -Server o -ConnectionString.' }
        if ($null -eq $Credencial) { $Credencial = Read-CredencialPopup -Titulo 'Conexión a Sybase' -Usuario $Usuario }
        $cs += 'UID={0};PWD={1};' -f $Credencial.UserName, $Credencial.GetNetworkCredential().Password
    }
    else { $cs = $ConnectionString }

    if (-not (Test-Path -LiteralPath $Carpeta)) { $null = New-Item -ItemType Directory -Path $Carpeta -Force }
    $rutaPend = Join-Path $Carpeta 'Debitos_Pendientes.csv'
    $rutaLoop = Join-Path $Carpeta 'Movimientos_Loop.csv'

    Write-Log "Paso 1: conectando a Sybase (débitos pendientes: todos; movimientos de Loop: $diaTxt)..."
    $conn = New-Object System.Data.Odbc.OdbcConnection($cs)
    try {
        $conn.Open()
        $f1 = Export-ConsultaCsv -Conexion $conn -Sql $sqlPendientes -Ruta $rutaPend -TimeoutSeg $TimeoutSeg
        if ($f1 -eq 0) { Write-Log 'Débitos pendientes: sin filas (CSV solo con el encabezado).' 'WARN' }
        Write-Log ('Débitos pendientes: {0} fila(s) en {1}' -f $f1, $rutaPend) 'OK'
        $f2 = Export-ConsultaCsv -Conexion $conn -Sql $sqlLoop -Ruta $rutaLoop -TimeoutSeg $TimeoutSeg
        if ($f2 -eq 0) { Write-Log 'Movimientos de Loop: sin filas (CSV solo con el encabezado).' 'WARN' }
        Write-Log ('Movimientos de Loop: {0} fila(s) en {1}' -f $f2, $rutaLoop) 'OK'
    }
    finally { $conn.Dispose() }
    return [pscustomobject]@{ Loop = $rutaLoop; Pendientes = $rutaPend }
}

# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------
try {
    if ($SoloGeneral) {
        Write-Log 'Modo SoloGeneral: se reconstruye el dashboard general desde los HTML diarios.'
        $null = Export-DashboardGeneral -OutputDir (Resolve-Path -LiteralPath $OutputDir).Path
        exit 0
    }
    if ($Descargar) {
        if (-not (Test-Path -LiteralPath $OutputDir)) { $null = New-Item -ItemType Directory -Path $OutputDir -Force }
        if (-not $DescargaDir) { $DescargaDir = Join-Path (Resolve-Path -LiteralPath $OutputDir).Path 'Entrada' }
        $fechaDia = $null
        if ($PSBoundParameters.ContainsKey('Fecha')) { $fechaDia = $Fecha }
        $bajados = Invoke-Descarga -Carpeta $DescargaDir -FechaDia $fechaDia -Dsn $Dsn -ConnectionString $ConnectionString -Server $Server -Port $Port -Driver $Driver -Usuario $Usuario -Credencial $Credencial -TimeoutSeg $Timeout
        if ($null -eq $bajados) { exit 0 }
        $LoopPath = $bajados.Loop
        $PendientesPath = $bajados.Pendientes
        Write-Log 'Paso 2: conciliación.'
    }
    Write-Log 'Inicio de la conciliación.'
    $loop = Read-Loop -Path $LoopPath
    $base = $null
    if ($LoopBasePath) { $base = Read-Loop -Path $LoopBasePath }
    $loopFinal = Merge-Loop -Loop $loop -Base $base
    $pend = Read-Pendientes -Path $PendientesPath
    $trans = Read-Transferencias -Path $TransferenciasPath

    $conc = Get-Conciliacion -Loop $loopFinal -Pendientes $pend -Transfers $trans
    $exc = Get-Excepciones -Conc $conc
    $ppd = Get-PendientesPorDia -Conc $conc

    foreach ($m in 0, 2) {
        $t = $conc.Totales[$m]
        Write-Log ('{0}: neto {1}; pendientes {2}; sin cubrir {3}.' -f (Get-NombreMoneda $m), (Format-Importe $t.Neto $m), (Format-Importe $t.Pendientes $m), (Format-Importe $t.SinCubrir $m))
    }
    foreach ($m in @('OK', 'Débito pendiente', 'Una partida', 'Importes distintos', 'Más de 2 partidas')) {
        Write-Log ('  {0}: {1}' -f $m, @($conc.Comprobantes | Where-Object { $_.Marca -eq $m }).Count)
    }
    $x = $conc.Cruce
    Write-Log ("Transfers: {0} operaciones, {1} en Loop, {2} FAILED sin registro, {3} APPROVED sin registro, {4} APPROVED pendientes de contabilizar, {5} fuera del período." -f $x.Total, $x.EnLoop, $x.FailedSinRegistro, $x.ApprovedSinRegistro, $x.ApprovedPendientes, $x.ApprovedFuera)

    if (-not (Test-Path -LiteralPath $OutputDir)) { $null = New-Item -ItemType Directory -Path $OutputDir -Force }
    $outDir = (Resolve-Path -LiteralPath $OutputDir).Path
    $dirDia = Join-Path $outDir 'Diario'
    if (-not (Test-Path -LiteralPath $dirDia)) { $null = New-Item -ItemType Directory -Path $dirDia -Force }
    $meta = [pscustomobject]@{
        LoopArchivo = $loop.Archivo; BaseArchivo = $(if ($base) { $base.Archivo } else { '' }); Faltantes = $loopFinal.Faltantes
        PendArchivo = $pend.Archivo; TransArchivo = $trans.Archivo; Dia = $null
    }

    # Un Excel y un HTML por día
    $fechas = @()
    $fechas += @($conc.Comprobantes | ForEach-Object { $_.Fecha })
    $fechas += @($conc.Hallazgos | Where-Object { $null -ne $_.Fecha } | ForEach-Object { ([datetime]$_.Fecha).Date })
    $dias = @($fechas | Sort-Object -Unique)
    Write-Log "Se generan $($dias.Count) reportes diarios en $dirDia"
    foreach ($dia in $dias) {
        $sd = Get-ConciliacionDia -Conc $conc -Dia $dia
        $ed = Get-Excepciones -Conc $sd -Silencioso
        $pd = Get-PendientesPorDia -Conc $sd
        $metaDia = [pscustomobject]@{
            LoopArchivo = $meta.LoopArchivo; BaseArchivo = $meta.BaseArchivo; Faltantes = $meta.Faltantes
            PendArchivo = $meta.PendArchivo; TransArchivo = $meta.TransArchivo; Dia = $dia
        }
        $baseDia = Join-Path $dirDia ('Conciliacion_Loop_' + $dia.ToString('yyyyMMdd', $script:Inv))
        $nExc = @($ed | Where-Object { $_.EsExcepcion }).Count
        Write-Log ('{0}: {1} comprobantes, {2} excepciones.' -f (Format-Fecha $dia), @($sd.Comprobantes).Count, $nExc)
        $null = Export-ExcelReport -Conc $sd -Excepciones $ed -PendDia $pd -Meta $metaDia -Base $baseDia
        $null = Export-HtmlDashboard -Conc $sd -Excepciones $ed -PendDia $pd -Meta $metaDia -Base $baseDia -EmbedirDatos
    }

    if ($IncluirPeriodo) {
        $baseOut = Join-Path $outDir ('Conciliacion_Loop_Periodo_' + (Get-Date).ToString('yyyyMMdd_HHmm', $script:Inv))
        $null = Export-ExcelReport -Conc $conc -Excepciones $exc -PendDia $ppd -Meta $meta -Base $baseOut
        $null = Export-HtmlDashboard -Conc $conc -Excepciones $exc -PendDia $ppd -Meta $meta -Base $baseOut
    }

    # Dashboard general: se alimenta de los HTML diarios (incluidos los de corridas anteriores)
    $null = Export-DashboardGeneral -OutputDir $outDir
    Write-Log 'Conciliación finalizada.' 'OK'
    exit 0
}
catch {
    Write-Log ('Error: ' + $_.Exception.Message) 'ERROR'
    if ($_.InvocationInfo) { Write-Log ('En: ' + $_.InvocationInfo.PositionMessage.Replace("`n", ' ')) 'ERROR' }
    exit 1
}
