<#
.SYNOPSIS
    Ejecuta en SAP Sybase (ASE) las consultas de débitos pendientes y de movimientos de Loop, y guarda cada resultado en un CSV.

.DESCRIPTION
    Compatible con Windows PowerShell 5.1 y PowerShell 7. Usa ODBC (System.Data.Odbc), por lo que
    necesita el driver ODBC de Sybase/SAP ASE instalado con la misma arquitectura (32 o 64 bits)
    que el PowerShell que se usa. Trae los débitos pendientes (por fecorigen) y los movimientos de Loop (por fecmov)
    del día anterior (si hoy es lunes, desde el viernes hasta el domingo), o del día indicado en -Fecha. El CSVs quedan listos para usar como -PendientesPath y -LoopPath en
    Conciliacion-Loop.ps1 (fechas d/M/yyyy HH:mm:ss.fff, decimales con punto, UTF-8).

.PARAMETER Dsn               Nombre del DSN ODBC ya configurado (opción más simple).
.PARAMETER ConnectionString  Cadena de conexión ODBC completa (alternativa a -Dsn / -Server).
.PARAMETER Server            Servidor ASE (si no se usa -Dsn ni -ConnectionString).
.PARAMETER Port              Puerto del servidor ASE (por defecto 5000).
.PARAMETER Driver            Nombre del driver ODBC (por defecto 'Adaptive Server Enterprise').
.PARAMETER Usuario           Usuario de la base. Si se indica, la ventana emergente pide solo la clave.
.PARAMETER Credencial        Credencial de la base. Si no se pasa, se pide la clave (y el usuario si falta) en una ventana emergente.
.PARAMETER Fecha             Día a consultar (por defecto, ayer).
.PARAMETER OutputPath        Ruta del CSV de débitos pendientes (por defecto .\Debitos_Pendientes.csv).
.PARAMETER LoopOutputPath    Ruta del CSV de movimientos de Loop (por defecto Movimientos_Loop.csv, junto al de pendientes).
.PARAMETER Timeout           Tiempo máximo de la consulta en segundos (por defecto 300).

.EXAMPLE
    .\Exportar-Pendientes.ps1 -Dsn BANKSYS -OutputPath .\Debitos_Pendientes.csv

.EXAMPLE
    .\Exportar-Pendientes.ps1 -Server srvase01 -Port 5000 -Fecha 2026-10-22
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Dsn,
    [string]$ConnectionString,
    [string]$Server,
    [int]$Port = 5000,
    [string]$Driver = 'Adaptive Server Enterprise',
    [string]$Usuario,
    [System.Management.Automation.PSCredential]$Credencial,
    [datetime]$Fecha,
    [string]$OutputPath = (Join-Path (Get-Location).Path 'Debitos_Pendientes.csv'),
    [string]$LoopOutputPath,
    [int]$Timeout = 300
)

$ErrorActionPreference = 'Stop'
$inv = [System.Globalization.CultureInfo]::InvariantCulture
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

function Write-Log {
    param([string]$Mensaje, [string]$Nivel = 'INFO')
    $color = 'Gray'
    if ($Nivel -eq 'WARN') { $color = 'Yellow' } elseif ($Nivel -eq 'ERROR') { $color = 'Red' } elseif ($Nivel -eq 'OK') { $color = 'Green' }
    Write-Host ('[{0}] {1}' -f $Nivel, $Mensaje) -ForegroundColor $color
}

function Format-CampoCsv {
    # Valor en texto invariante, sin saltos de línea y con comillas si hace falta
    param($Valor)
    if ($null -eq $Valor -or $Valor -is [System.DBNull]) { return '' }
    if ($Valor -is [datetime]) { $t = $Valor.ToString('d/M/yyyy HH:mm:ss.fff', $inv) }
    elseif ($Valor -is [decimal] -or $Valor -is [double] -or $Valor -is [single]) { $t = ([System.IFormattable]$Valor).ToString($null, $inv) }
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

$conn = $null
try {
    # Mismo criterio de fechas para las dos consultas: ayer por defecto; si hoy es lunes, desde el viernes (@hoy - 3)
    if ($PSBoundParameters.ContainsKey('Fecha')) {
        $desde = $Fecha.Date
        $hasta = $desde.AddDays(1)
    }
    else {
        $hasta = (Get-Date).Date
        if ($hasta.DayOfWeek -eq [System.DayOfWeek]::Monday) { $desde = $hasta.AddDays(-3) }
        else { $desde = $hasta.AddDays(-1) }
    }
    $fechasSql = "declare @hoy datetime, @ayer datetime`r`nselect @hoy  = '{0}'`r`nselect @ayer = '{1}'`r`n" -f $hasta.ToString('yyyyMMdd', $inv), $desde.ToString('yyyyMMdd', $inv)
    $ultimo = $hasta.AddDays(-1)
    if ($desde -eq $ultimo) { $diaTxt = $desde.ToString('dd/MM/yyyy', $inv) }
    else { $diaTxt = 'del {0} al {1}' -f $desde.ToString('dd/MM/yyyy', $inv), $ultimo.ToString('dd/MM/yyyy', $inv) }

    $sqlPendientes = $fechasSql + @"

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
and b.fecorigen >= @ayer
and b.fecorigen <  @hoy

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

    # Cadena de conexión
    if (-not $ConnectionString) {
        if ($Dsn) { $cs = "DSN=$Dsn;" }
        elseif ($Server) { $cs = "Driver={$Driver};Server=$Server;Port=$Port;" }
        else { throw 'Indicá -Dsn, -Server o -ConnectionString para conectarte a la base.' }
        if ($null -eq $Credencial) { $Credencial = Read-CredencialPopup -Titulo 'Conexión a Sybase' -Usuario $Usuario }
        $cs += 'UID={0};PWD={1};' -f $Credencial.UserName, $Credencial.GetNetworkCredential().Password
    }
    else { $cs = $ConnectionString }

    if (-not $LoopOutputPath) {
        $carpeta = Split-Path -Parent $OutputPath
        if (-not $carpeta) { $carpeta = (Get-Location).Path }
        $LoopOutputPath = Join-Path $carpeta 'Movimientos_Loop.csv'
    }

    Write-Log "Conectando a Sybase para traer los datos: $diaTxt..."
    $conn = New-Object System.Data.Odbc.OdbcConnection($cs)
    $conn.Open()

    $f1 = Export-ConsultaCsv -Conexion $conn -Sql $sqlPendientes -Ruta $OutputPath -TimeoutSeg $Timeout
    if ($f1 -eq 0) { Write-Log 'Débitos pendientes: sin filas; se guardó el CSV solo con el encabezado.' 'WARN' }
    Write-Log ('Débitos pendientes: {0} fila(s) guardadas en {1}' -f $f1, $OutputPath) 'OK'

    $f2 = Export-ConsultaCsv -Conexion $conn -Sql $sqlLoop -Ruta $LoopOutputPath -TimeoutSeg $Timeout
    if ($f2 -eq 0) { Write-Log 'Movimientos de Loop: sin filas; se guardó el CSV solo con el encabezado.' 'WARN' }
    Write-Log ('Movimientos de Loop: {0} fila(s) guardadas en {1}' -f $f2, $LoopOutputPath) 'OK'
    exit 0
}
catch {
    Write-Log ('Error: ' + $_.Exception.Message) 'ERROR'
    exit 1
}
finally {
    if ($null -ne $conn) { $conn.Dispose() }
}
