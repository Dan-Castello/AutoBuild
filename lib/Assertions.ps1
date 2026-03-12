#Requires -Version 5.1
# lib/Assertions.ps1
# Funciones de validacion y prerequisitos para tareas.

Set-StrictMode -Version Latest

function Assert-Param {
    <#
    .SYNOPSIS
        Valida que un parametro requerido no sea nulo ni vacio.
    #>
    param(
        [string]$Name,
        [string]$Value,
        [hashtable]$Context
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        $msg = "Parametro requerido no especificado: $Name"
        if ($null -ne $Context) {
            Write-BuildLog $Context 'ERROR' $msg
        }
        throw $msg
    }
}

function Assert-FileExists {
    <#
    .SYNOPSIS
        Valida que un archivo exista antes de continuar.
    #>
    param(
        [string]$Path,
        [string]$Label = '',
        [hashtable]$Context
    )

    if (-not (Test-Path $Path)) {
        $msg = "Archivo no encontrado$(if ($Label) {" ($Label)"}): $Path"
        if ($null -ne $Context) {
            Write-BuildLog $Context 'ERROR' $msg
        }
        throw $msg
    }
}

function Assert-SapSession {
    <#
    .SYNOPSIS
        Valida que se haya obtenido una sesion SAP valida.
    #>
    param(
        $Session,
        [hashtable]$Context
    )

    if ($null -eq $Session) {
        $msg = 'No hay sesion SAP activa. Iniciar SAP GUI y conectarse al sistema.'
        if ($null -ne $Context) {
            Write-BuildLog $Context 'ERROR' $msg
        }
        throw $msg
    }
}

function Test-TaskAsset {
    <#
    .SYNOPSIS
        Comprueba multiples prerequisitos (archivos + parametros) de una sola vez.
        Lanza excepcion en el primer fallo encontrado.
    .PARAMETER Files
        Hashtable nombre->ruta de archivos que deben existir.
    .PARAMETER Params
        Hashtable nombre->valor de parametros que deben ser no vacios.
    #>
    param(
        [hashtable]$Context,
        [hashtable]$Files  = @{},
        [hashtable]$Params = @{}
    )

    foreach ($kv in $Files.GetEnumerator()) {
        Assert-FileExists -Path $kv.Value -Label $kv.Key -Context $Context
    }

    foreach ($kv in $Params.GetEnumerator()) {
        Assert-Param -Name $kv.Key -Value $kv.Value -Context $Context
    }
}
