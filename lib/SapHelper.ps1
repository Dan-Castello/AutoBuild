#Requires -Version 5.1
# lib/SapHelper.ps1
# Ayudas para SAP GUI Scripting en PS 5.1.
# SAP GUI 800 Final release 8000.1.11.1161
#
# REGLA COM: Todo objeto COM intermedio debe capturarse en variable y
# liberarse en un bloque finally. Esta regla es identica a la de
# ExcelHelper.ps1 y WordHelper.ps1.
#
# PATRON en Get-SapSession:
#   $rot -> $sapgui -> $app -> $conn -> $sess
#   Solo $sess se retorna al llamante.
#   $rot, $sapgui, $app, $conn se liberan en finally antes de retornar.
#   El llamante libera $sess con Release-SapSession cuando termina.

Set-StrictMode -Version Latest

function Get-SapSession {
    <#
    .SYNOPSIS
        Obtiene la sesion SAP GUI activa con gestion correcta del ciclo
        de vida COM de todos los objetos intermedios.
    .NOTES
        FIX-SAP-COM-01: $rot, $sapgui, $app y $conn son objetos COM con
        ref-count propio. Sin release explicito mantienen handles abiertos
        en el proceso saplogon.exe, agotando slots de scripting engine en
        sesiones largas.
        El objeto $sess se retorna al llamante. El llamante DEBE liberarlo
        llamando Release-SapSession al terminar.
    .OUTPUTS
        Objeto GuiSession o $null si no hay sesion activa o hay error.
    #>
    param(
        [hashtable]$Context,
        [int]$ConnectionIndex = 0,
        [int]$SessionIndex    = 0
    )

    Write-BuildLog $Context 'DEBUG' 'Buscando sesion SAP GUI activa'

    $rot    = $null
    $sapgui = $null
    $app    = $null
    $conn   = $null

    try {
        $rot = New-Object -ComObject 'SapROTWr.SapROTWrapper' -ErrorAction Stop

        $sapgui = $rot.GetROTEntry('SAPGUI')
        if ($null -eq $sapgui) {
            Write-BuildLog $Context 'WARN' 'No hay instancia de SAP GUI en ejecucion'
            return $null
        }

        $app = $sapgui.GetScriptingEngine()
        if ($null -eq $app) {
            Write-BuildLog $Context 'WARN' 'SAP GUI no expone scripting engine'
            return $null
        }

        $conn = $app.Children($ConnectionIndex)
        if ($null -eq $conn) {
            Write-BuildLog $Context 'WARN' "No hay conexion SAP en indice $ConnectionIndex"
            return $null
        }

        $sess = $conn.Children($SessionIndex)
        if ($null -eq $sess) {
            Write-BuildLog $Context 'WARN' "No hay sesion SAP en indice $SessionIndex"
            return $null
        }

        # Leer informacion de sesion antes de liberar intermedios.
        # $sess.Info es un subcomponente del objeto sesion retornado,
        # no un intermedio de adquisicion de la cadena ROT.
        $sysName  = ''
        $userName = ''
        try {
            $sessInfo = $sess.Info
            $sysName  = $sessInfo.SystemName
            $userName = $sessInfo.User
            Release-ComObject $sessInfo
        } catch { }

        Write-BuildLog $Context 'INFO' "Sesion SAP obtenida: $sysName / $userName"
        return $sess

    } catch {
        Write-BuildLog $Context 'ERROR' "Error al obtener sesion SAP: $_"
        return $null

    } finally {
        # Liberar intermedios en orden inverso. La clausula finally se
        # ejecuta incluso cuando hay un return explicito en PS 5.1.
        # $sess NO se libera aqui: ha sido retornado al llamante.
        if ($null -ne $conn)   { Release-ComObject $conn;   $conn   = $null }
        if ($null -ne $app)    { Release-ComObject $app;    $app    = $null }
        if ($null -ne $sapgui) { Release-ComObject $sapgui; $sapgui = $null }
        if ($null -ne $rot)    { Release-ComObject $rot;    $rot    = $null }
    }
}

function Release-SapSession {
    <#
    .SYNOPSIS
        Libera el objeto de sesion SAP retornado por Get-SapSession.
        Llamar en el bloque finally de cualquier tarea que use SAP.
    #>
    param($Session)

    if ($null -eq $Session) { return }
    try {
        Release-ComObject $Session
    } catch { }
}

function Invoke-SapTransaction {
    <#
    .SYNOPSIS
        Abre una transaccion SAP en la sesion activa.
    #>
    param(
        [hashtable]$Context,
        $Session,
        [string]$TCode
    )

    if ($null -eq $Session) { throw 'Sesion SAP nula' }
    Write-BuildLog $Context 'INFO' "Ejecutando transaccion: $TCode"
    try {
        $Session.StartTransaction($TCode)
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al ejecutar transaccion $TCode : $_"
        throw
    }
}

function Set-SapField {
    <#
    .SYNOPSIS
        Establece el valor de un campo SAP por su ID de componente.
    #>
    param(
        [hashtable]$Context,
        $Session,
        [string]$FieldId,
        [string]$Value
    )

    if ($null -eq $Session) { throw 'Sesion SAP nula' }
    $field = $null
    try {
        $field = $Session.FindById($FieldId)
        $field.Text = $Value
    } catch {
        Write-BuildLog $Context 'WARN' "No se pudo establecer campo $FieldId : $_"
        throw
    } finally {
        if ($null -ne $field) { Release-ComObject $field; $field = $null }
    }
}

function Get-SapField {
    <#
    .SYNOPSIS
        Obtiene el valor de un campo SAP por su ID de componente.
    #>
    param(
        [hashtable]$Context,
        $Session,
        [string]$FieldId
    )

    if ($null -eq $Session) { throw 'Sesion SAP nula' }
    $field = $null
    try {
        $field = $Session.FindById($FieldId)
        return $field.Text
    } catch {
        Write-BuildLog $Context 'WARN' "No se pudo leer campo $FieldId : $_"
        return $null
    } finally {
        if ($null -ne $field) { Release-ComObject $field; $field = $null }
    }
}

function Invoke-SapButton {
    <#
    .SYNOPSIS
        Presiona un boton SAP por su ID de componente.
    #>
    param(
        [hashtable]$Context,
        $Session,
        [string]$ButtonId
    )

    if ($null -eq $Session) { throw 'Sesion SAP nula' }
    $btn = $null
    try {
        $btn = $Session.FindById($ButtonId)
        $btn.Press()
        Write-BuildLog $Context 'DEBUG' "Boton presionado: $ButtonId"
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al presionar boton $ButtonId : $_"
        throw
    } finally {
        if ($null -ne $btn) { Release-ComObject $btn; $btn = $null }
    }
}

function Export-SapTableToArray {
    <#
    .SYNOPSIS
        Exporta una tabla SAP (GuiTableControl o ALV) a un array de hashtables.
    .OUTPUTS
        Array de hashtables, una por fila.
    #>
    param(
        [hashtable]$Context,
        $Session,
        [string]$TableId,
        [string[]]$Columns
    )

    if ($null -eq $Session) { throw 'Sesion SAP nula' }

    $rows  = @()
    $table = $null
    try {
        $table    = $Session.FindById($TableId)
        $rowCount = $table.RowCount

        Write-BuildLog $Context 'INFO' "Exportando tabla SAP: $rowCount filas, $($Columns.Count) columnas"

        for ($r = 0; $r -lt $rowCount; $r++) {
            $row = @{}
            foreach ($col in $Columns) {
                $cell = $null
                try {
                    $cell      = $table.GetCell($r, $col)
                    $row[$col] = $cell.Text
                } catch {
                    $row[$col] = ''
                } finally {
                    if ($null -ne $cell) { Release-ComObject $cell; $cell = $null }
                }
            }
            $rows += $row
        }
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al exportar tabla $TableId : $_"
        throw
    } finally {
        if ($null -ne $table) { Release-ComObject $table; $table = $null }
    }

    return $rows
}
