#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_entorno.ps1
# @Description : Diagnostico del entorno PowerShell y sistema operativo
# @Category    : Utilidad
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1. Sin COM. Sin parametros requeridos.

# Synopsis: Verifica version PS, politica de ejecucion, OS y carpetas del proyecto
task diag_entorno {
    $ctx = New-TaskContext `
        -TaskName 'diag_entorno' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico de entorno'

    $errores  = 0
    $avisos   = 0
    $Root     = Split-Path $BuildRoot -Parent

    # ---- PowerShell --------------------------------------------------------
    $psv = $PSVersionTable.PSVersion
    Write-Build Cyan  "  [PS]  Version        : $($psv.Major).$($psv.Minor).$($psv.Build)"
    if ($psv.Major -eq 5 -and $psv.Minor -eq 1) {
        Write-Build Green "  [PS]  Version         : OK (5.1)"
    } else {
        Write-Build Yellow "  [PS]  WARN: se recomienda PS 5.1 para compatibilidad COM"
        $avisos++
    }

    $edition = $PSVersionTable.PSEdition
    Write-Build Cyan  "  [PS]  Edition         : $edition"
    if ($edition -eq 'Desktop') {
        Write-Build Green "  [PS]  Edition          : OK (Desktop)"
    } else {
        Write-Build Yellow "  [PS]  WARN: Edition '$edition' puede tener limitaciones COM"
        $avisos++
    }

    # ---- Politica de ejecucion ---------------------------------------------
    $pol = Get-ExecutionPolicy -Scope CurrentUser
    Write-Build Cyan  "  [PS]  ExecPolicy       : $pol"
    if ($pol -in @('RemoteSigned','Unrestricted','Bypass')) {
        Write-Build Green "  [PS]  ExecPolicy        : OK"
    } else {
        Write-Build Red   "  [PS]  ERROR: ExecutionPolicy '$pol' puede bloquear scripts"
        Write-Build Red   "         Ejecutar: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser"
        $errores++
    }

    # ---- Sistema operativo -------------------------------------------------
    $os = (Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue)
    if ($null -ne $os) {
        Write-Build Cyan  "  [OS]  Caption          : $($os.Caption)"
        Write-Build Cyan  "  [OS]  Build             : $($os.BuildNumber)"
        Write-Build Cyan  "  [OS]  Arquitectura      : $($os.OSArchitecture)"
    } else {
        Write-Build Yellow "  [OS]  WARN: no se pudo obtener info del SO via WMI"
        $avisos++
    }

    # ---- .NET Framework ----------------------------------------------------
    try {
        $dotnet = (Get-ItemProperty `
            'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' `
            -ErrorAction Stop).Release
        $dotnetVer = switch ($dotnet) {
            { $_ -ge 533320 } { '4.8.1+' }
            { $_ -ge 528040 } { '4.8'    }
            { $_ -ge 461808 } { '4.7.2'  }
            { $_ -ge 460798 } { '4.7'    }
            default           { "release=$dotnet" }
        }
        Write-Build Cyan  "  [.NET] Framework       : $dotnetVer"
        Write-Build Green "  [.NET] Framework        : OK"
    } catch {
        Write-Build Yellow "  [.NET] WARN: no se pudo verificar .NET Framework"
        $avisos++
    }

    # ---- Carpetas del proyecto ---------------------------------------------
    foreach ($dir in @('input','output','reports','logs','tasks','lib','engine')) {
        $p = Join-Path $Root $dir
        if (Test-Path $p) {
            Write-Build Green "  [DIR] $dir : OK"
        } else {
            Write-Build Red   "  [DIR] $dir : FALTA"
            $errores++
        }
    }

    # ---- Invoke-Build portable ---------------------------------------------
    $ibPath = Join-Path $Root 'tools\InvokeBuild\Invoke-Build.ps1'
    if (Test-Path $ibPath) {
        Write-Build Green "  [IB]  Invoke-Build      : OK (portable)"
    } else {
        Write-Build Red   "  [IB]  Invoke-Build      : FALTA en tools\InvokeBuild\"
        $errores++
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan  ""
    if ($errores -eq 0 -and $avisos -eq 0) {
        Write-Build Green "  RESULTADO: OK - entorno listo"
    } elseif ($errores -eq 0) {
        Write-Build Yellow "  RESULTADO: $avisos aviso(s) - revisar antes de produccion"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es), $avisos aviso(s) - corregir antes de continuar"
    }

    Write-BuildLog $ctx 'INFO' "diag_entorno completado. Errores=$errores Avisos=$avisos"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        throw "Diagnostico detecto $errores error(es) en el entorno"
    }
}
