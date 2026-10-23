<#
.SYNOPSIS
    Concilia las transferencias QNET contra las partidas de Loop y los débitos
    pendientes. Genera un Excel (o CSV de respaldo) y un dashboard HTML.

.DESCRIPTION
    Compatible con Windows PowerShell 5.1 y PowerShell 7.
    Entradas: Loop (partidas), Débitos pendientes y Transferencias QNET (CSV).
    Salidas: Conciliacion_Loop_<yyyyMMdd_HHmm>.xlsx y .html en -OutputDir.
    Código de salida: 0 si no hubo excepciones de ejecución, 1 si las hubo.
    No se vuelcan nombres, CUIT ni CBU en las salidas.

.PARAMETER LoopPath            CSV de Loop (obligatorio).
.PARAMETER PendientesPath      CSV de débitos pendientes (obligatorio).
.PARAMETER TransferenciasPath  CSV de transferencias QNET (obligatorio).
.PARAMETER LoopBasePath        CSV base de Loop para completar el período anterior (opcional).
.PARAMETER OutputDir           Carpeta de salida (por defecto, la carpeta actual).

.EXAMPLE
    .\Conciliacion-Loop.ps1 -LoopPath .\SQL14.csv -LoopBasePath .\Tr_Loop_116261.csv `
        -PendientesPath .\SQL15.csv `
        -TransferenciasPath .\transferencias-qnet-2026-10-01_2026-10-23.csv `
        -OutputDir .\salida
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$LoopPath,
    [Parameter(Mandatory = $true)][string]$PendientesPath,
    [Parameter(Mandatory = $true)][string]$TransferenciasPath,
    [string]$LoopBasePath,
    [string]$OutputDir = (Get-Location).Path
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constantes y formato
# ---------------------------------------------------------------------------
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture
$script:Tol = [decimal]0.005
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
    $enLoop = 0; $failSin = 0; $apSin = 0; $apPend = 0; $apFuera = 0; $cbuMal = 0
    foreach ($t in $Transfers.Rows) {
        if ($t.CbuOk -eq $false) {
            $cbuMal++
            $esp = '02'; if ($t.Mon -eq 2) { $esp = '06' }
            $hall.Add((New-Hallazgo 'CBU origen no corresponde a la moneda' $t.Comnro $t.Cuenta $t.Mon $t.Fecha $t.Monto $t.Estado $null ("Dígitos 9 y 10 del CBU origen = '{0}', se esperaba '{1}'" -f $t.TipoCta, $esp)))
        }
        if ($porComnro.ContainsKey($t.Comnro)) { $enLoop++; continue }
        if ($t.Estado -eq 'FAILED') { $failSin++; continue }
        if ($null -eq $t.Fecha) { $apSin++; $hall.Add((New-Hallazgo 'APPROVED sin registro' $t.Comnro $t.Cuenta $t.Mon $null $t.Monto 'APPROVED' $null 'APPROVED sin fecha y sin asiento en Loop')); continue }
        if ($t.Fecha -ge $loopMax) { $apPend++ }
        elseif ($t.Fecha -lt $loopMin) { $apFuera++ }
        else {
            $apSin++
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
        Comprobantes = $ordenados; Hallazgos = $hall; Totales = $tot
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
    param($Conc)
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
    Write-Log "Excepciones detectadas: $n (más $(@($sorted).Count - $n) comprobantes con débito pendiente listados)."
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
    $bloques.Add(@{ Titulo = 'Cruce con transfers'; Cols = @('Concepto', 'Cantidad'); Formula = $false; Filas = @(
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
                @('Período de Loop', ('{0} al {1}' -f (Format-Fecha $Conc.LoopMin), (Format-Fecha $Conc.LoopMax))),
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

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log 'No está el módulo ImportExcel. Instalalo con: Install-Module ImportExcel -Scope CurrentUser' 'WARN'
        Write-Log 'Como respaldo se exportan las hojas como CSV.' 'WARN'
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
            'Comprobantes'      = @($filasComp | ForEach-Object { $_.'Fecha débito' = (& $fmtFecha $_.'Fecha débito'); $_.'Fecha crédito' = (& $fmtFecha $_.'Fecha crédito'); $_ })
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
            $null = Export-Excel -ExcelPackage $pkg -WorksheetName 'Comprobantes' -InputObject $filasComp -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
            $wc = $pkg.Workbook.Worksheets['Comprobantes']
            $uc = $filasComp.Count + 1
            $wc.Cells[2, 5, $uc, 7].Style.Numberformat.Format = '#,##0.00'
            $wc.Cells[2, 8, $uc, 9].Style.Numberformat.Format = 'dd/mm/yyyy'
            $wc.Cells[2, 3, $uc, 3].Style.Numberformat.Format = '0'
            $wc.Cells[2, 10, $uc, 10].Style.Numberformat.Format = '0'
            try { $wc.Cells[$wc.Dimension.Address].AutoFitColumns() } catch { }
        }
        else { $null = $pkg.Workbook.Worksheets.Add('Comprobantes') }

        # Resumen
        $ws = $pkg.Workbook.Worksheets.Add('Resumen')
        $ws.Cells[1, 1].Value = 'Conciliación Loop - Transferencias QNET'
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
        [void]$sb.Append("<line class=""grid"" x1=""$ml"" x2=""$($W - $mr)"" y1=""$(& $f $y)"" y2=""$(& $f $y)""/>")
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
    param($Conc, $Excepciones, $PendDia, $Meta, [string]$Base)
    $ruta = $Base + '.html'
    $sb = New-Object System.Text.StringBuilder
    $ap = { param($t) [void]$sb.AppendLine($t) }

    $cnt = { param($m) @($Conc.Comprobantes | Where-Object { $_.Marca -eq $m }).Count }
    $nOk = & $cnt 'OK'; $nPend = & $cnt 'Débito pendiente'; $nUna = & $cnt 'Una partida'; $nDist = & $cnt 'Importes distintos'
    $nDos = @($Conc.Comprobantes | Where-Object { $_.Partidas -eq 2 }).Count
    $excs = @($Excepciones | Where-Object { $_.EsExcepcion })
    $tp = $Conc.Totales[0]; $tu = $Conc.Totales[2]
    $x = $Conc.Cruce
    $periodo = '{0} al {1}' -f (Format-Fecha $Conc.LoopMin), (Format-Fecha $Conc.LoopMax)
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
    & $ap @'
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
svg{width:100%;height:auto;display:block}svg .grid{stroke:var(--line);stroke-width:1}svg .ax{fill:var(--muted);font-size:11px}
.s0{fill:var(--s0)}.s1{fill:var(--s1)}.s2{fill:var(--s2)}.s3{fill:var(--s3)}
.legend{display:flex;flex-wrap:wrap;gap:6px 16px;margin-bottom:6px}.sw{display:inline-block;width:11px;height:11px;border-radius:2px;margin-right:5px;vertical-align:-1px}
i.s0{background:var(--s0)}i.s1{background:var(--s1)}i.s2{background:var(--s2)}i.s3{background:var(--s3)}
#tip{position:fixed;display:none;pointer-events:none;background:#111;color:#fff;padding:4px 8px;border-radius:4px;font-size:12px;z-index:9}
@media (max-width:600px){h1{font-size:18px}.tile .val{font-size:17px}}
</style></head><body><main>
'@
    & $ap '<h1>Conciliación Loop - Transferencias QNET</h1>'
    & $ap ('<div class="meta"><span>Período de Loop: <b>{0}</b></span><span>Corte de transfers: <b>{1}</b></span><span>Generado: {2}</span></div>' -f (Get-Enc $periodo), (Get-Enc $corteTxt), (Get-Enc (Get-Date).ToString('dd/MM/yyyy HH:mm', $script:Inv)))
    & $ap ('<div class="resumen">{0}</div>' -f (Get-Enc $frase))

    & $ap '<h2>Comprobantes</h2><div class="grid">'
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

    & $ap '<h2>Cruce con transfers</h2><div class="grid">'
    & $ap (& $tile 'Operaciones' $x.Total $false '')
    & $ap (& $tile 'Encontradas en Loop' $x.EnLoop $false '')
    & $ap (& $tile 'FAILED sin registro' $x.FailedSinRegistro $false '')
    & $ap (& $tile 'APPROVED sin registro' $x.ApprovedSinRegistro ($x.ApprovedSinRegistro -gt 0) '')
    & $ap (& $tile 'APPROVED pendientes de contabilizar' $x.ApprovedPendientes $false 'No son excepción')
    & $ap (& $tile 'APPROVED fuera del período' $x.ApprovedFuera $false '')
    & $ap '</div>'

    & $ap '<h2>Comprobantes por fecha</h2><div class="card">'
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

# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------
try {
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
    $stamp = (Get-Date).ToString('yyyyMMdd_HHmm', $script:Inv)
    $baseOut = Join-Path (Resolve-Path -LiteralPath $OutputDir).Path "Conciliacion_Loop_$stamp"
    $meta = [pscustomobject]@{
        LoopArchivo = $loop.Archivo; BaseArchivo = $(if ($base) { $base.Archivo } else { '' }); Faltantes = $loopFinal.Faltantes
        PendArchivo = $pend.Archivo; TransArchivo = $trans.Archivo
    }
    $null = Export-ExcelReport -Conc $conc -Excepciones $exc -PendDia $ppd -Meta $meta -Base $baseOut
    $null = Export-HtmlDashboard -Conc $conc -Excepciones $exc -PendDia $ppd -Meta $meta -Base $baseOut
    Write-Log 'Conciliación finalizada.' 'OK'
    exit 0
}
catch {
    Write-Log ('Error: ' + $_.Exception.Message) 'ERROR'
    if ($_.InvocationInfo) { Write-Log ('En: ' + $_.InvocationInfo.PositionMessage.Replace("`n", ' ')) 'ERROR' }
    exit 1
}
