<#
.SYNOPSIS
    Baja de SAP Sybase (ASE) los débitos pendientes y los movimientos de Loop y los guarda en CSV.

.DESCRIPTION
    Compatible con Windows PowerShell 5.1 y PowerShell 7. Usa ODBC (System.Data.Odbc): necesita el driver
    ODBC de Sybase/SAP ASE instalado con la misma arquitectura (32 o 64 bits) que el PowerShell.
    Solo lee (select). La clave se pide en una ventana emergente.
      - Débitos pendientes: todos. Se guardan en <CarpetaDP>\Debitos_Pendientes_<yyyyMMdd>.csv
      - Movimientos de Loop: desde el último día hábil anterior hasta ayer (lunes: desde el viernes).
        Se guardan en <CarpetaLoop>\Movimientos_Loop_<yyyyMMdd>.csv
    No se ejecuta en fines de semana ni feriados (lista 2026 en $script:Feriados); con -Fecha corre siempre.
    Las carpetas son las mismas que lee Conciliacion-Loop.ps1. Código de salida: 0 si terminó bien, 1 si hubo error.

.PARAMETER Fecha          Día puntual a bajar (dd/MM/yyyy o yyyy-MM-dd). Sin esto: desde el último día hábil anterior.
.PARAMETER FechaHasta     Con -Fecha, último día del rango a bajar (por defecto, el mismo día).
.PARAMETER Server         Servidor ASE (o -Dsn / -ConnectionString).
.PARAMETER Usuario        Usuario de Sybase; la clave se pide en una ventana emergente.
.PARAMETER Dsn            DSN ODBC de Sybase.
.PARAMETER Port           Puerto ASE (por defecto 5000).
.PARAMETER Driver         Driver ODBC (por defecto 'Adaptive Server Enterprise').
.PARAMETER ConnectionString  Cadena de conexión ODBC completa.
.PARAMETER Credencial     Credencial ya armada (evita la ventana).
.PARAMETER Timeout        Segundos máximos por consulta (por defecto 300).
.PARAMETER CarpetaLoop    Destino de los Loop (por defecto, 'entrada_loop' junto al script).
.PARAMETER CarpetaDP      Destino de los débitos pendientes (por defecto, 'entrada_DP' junto al script).

.EXAMPLE
    .\Descargar-Sybase.ps1

.EXAMPLE
    .\Descargar-Sybase.ps1 -Fecha 20/10/2026 -FechaHasta 23/10/2026
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    # ---- CONFIGURACIÓN ----
    [string]$Server = 'aconquija1',
    [int]$Port = 5000,
    [string]$Dsn = '',
    [string]$Driver = 'Adaptive Server Enterprise',
    [string]$Usuario = 'SRV_CONCILIACIONES',
    [string]$CarpetaLoop = (Join-Path $PSScriptRoot 'entrada_loop'),
    [string]$CarpetaDP = (Join-Path $PSScriptRoot 'entrada_DP'),
    # ---- Opcionales ----
    [string]$ConnectionString,
    [System.Management.Automation.PSCredential]$Credencial,
    [string]$Fecha,
    [string]$FechaHasta,
    [int]$Timeout = 300
)

$ErrorActionPreference = 'Stop'
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

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

function Write-Log {
    param([string]$Mensaje, [string]$Nivel = 'INFO')
    $color = 'Gray'
    if ($Nivel -eq 'WARN') { $color = 'Yellow' }
    elseif ($Nivel -eq 'ERROR') { $color = 'Red' }
    elseif ($Nivel -eq 'OK') { $color = 'Green' }
    Write-Host ('[{0}] {1}' -f $Nivel, $Mensaje) -ForegroundColor $color
}

function Format-Fecha {
    param($Fecha)
    if ($null -eq $Fecha) { return '' }
    return ([datetime]$Fecha).ToString('dd/MM/yyyy', $script:Inv)
}

function Convert-Fecha {
    param($Texto, [string[]]$Formatos)
    if ($null -eq $Texto) { return $null }
    $t = ([string]$Texto).Trim()
    if ($t -eq '') { return $null }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact($t, $Formatos, $script:Inv, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt }
    return $null
}

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
    param([string]$CarpetaLoop, [string]$CarpetaDP, $FechaDia, $FechaFin, [string]$Dsn, [string]$ConnectionString, [string]$Server, [int]$Port, [string]$Driver, [string]$Usuario, $Credencial, [int]$TimeoutSeg)
    if ($null -ne $FechaDia) {
        $desde = ([datetime]$FechaDia).Date
        $hasta = $desde.AddDays(1)
        if ($null -ne $FechaFin) {
            $fin = ([datetime]$FechaFin).Date
            if ($fin -lt $desde) { throw 'FechaHasta no puede ser anterior a Fecha.' }
            $hasta = $fin.AddDays(1)
        }
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
        else { throw 'Falta la conexión a Sybase: completá Server (o Dsn) en el bloque CONFIGURACIÓN al inicio del script, o pasá -Server / -Dsn / -ConnectionString.' }
        if ($null -eq $Credencial) { $Credencial = Read-CredencialPopup -Titulo 'Conexión a Sybase' -Usuario $Usuario }
        $cs += 'UID={0};PWD={1};' -f $Credencial.UserName, $Credencial.GetNetworkCredential().Password
    }
    else { $cs = $ConnectionString }

    foreach ($c in @($CarpetaLoop, $CarpetaDP)) { if (-not (Test-Path -LiteralPath $c)) { $null = New-Item -ItemType Directory -Path $c -Force } }
    # Archivos bajados con la fecha de la corrida en el nombre (queda el historial)
    $sufijo = (Get-Date).ToString('yyyyMMdd', $script:Inv)
    $rutaPend = Join-Path $CarpetaDP "Debitos_Pendientes_$sufijo.csv"
    $rutaLoop = Join-Path $CarpetaLoop "Movimientos_Loop_$sufijo.csv"

    Write-Log "Conectando a Sybase (débitos pendientes: todos; movimientos de Loop: $diaTxt)..."
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

try {
    $fechaDia = $null; $fechaFin = $null
    $fmtsFecha = @('dd/MM/yyyy', 'd/M/yyyy', 'yyyy-MM-dd')
    if ($Fecha) {
        $fechaDia = Convert-Fecha $Fecha $fmtsFecha
        if ($null -eq $fechaDia) { throw "Fecha inválida '$Fecha' (use dd/MM/yyyy o yyyy-MM-dd)." }
    }
    if ($FechaHasta) {
        if (-not $Fecha) { throw '-FechaHasta requiere -Fecha.' }
        $fechaFin = Convert-Fecha $FechaHasta $fmtsFecha
        if ($null -eq $fechaFin) { throw "FechaHasta inválida '$FechaHasta' (use dd/MM/yyyy o yyyy-MM-dd)." }
    }
    $bajados = Invoke-Descarga -CarpetaLoop $CarpetaLoop -CarpetaDP $CarpetaDP -FechaDia $fechaDia -FechaFin $fechaFin -Dsn $Dsn -ConnectionString $ConnectionString -Server $Server -Port $Port -Driver $Driver -Usuario $Usuario -Credencial $Credencial -TimeoutSeg $Timeout
    if ($null -ne $bajados) { Write-Log 'Descarga finalizada. Ahora podés ejecutar Conciliacion-Loop.ps1.' 'OK' }
    exit 0
}
catch {
    Write-Log ('Error: ' + $_.Exception.Message) 'ERROR'
    exit 1
}
