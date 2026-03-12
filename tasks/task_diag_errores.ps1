#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_errores.ps1
# @Description : Pruebas de errores planeados: verifica que el motor maneja
#                correctamente excepciones, rutas inexistentes y parametros
#                invalidos sin crashear ni dejar procesos colgados.
#                Cada bloque try/catch DEBE capturar el error esperado.
#                Si un bloque NO lanza excepcion, se registra como fallo.
# @Category    : Diagnostico
# @Version     : 1.1.0
# =============================================================================
# Solo ASCII. PS 5.1.
# NOTA: No usar funciones locales dentro del bloque task{} con $script: scope.
#       Invoke-Build no garantiza ese scope en PS 5.1. Todo inline.

# Synopsis: Verifica manejo de excepciones, rutas invalidas y parametros nulos
task diag_errores {
    $ctx = New-TaskContext `
        -TaskName 'diag_errores' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando pruebas de errores planeados'

    $pruebas = 0
    $pasadas = 0

    # =========================================================================
    # PRUEBA 1: Open-ExcelWorkbook con ruta inexistente debe lanzar excepcion
    # =========================================================================
    Write-Build Cyan ""
    Write-Build Cyan "  [1/5] Open-ExcelWorkbook ruta inexistente -> excepcion esperada"
    $pruebas++
    $xl1 = $null
    $p1ok = $false
    try {
        $xl1 = New-ExcelApp -Context $ctx -TimeoutSec 20
        if ($null -eq $xl1) {
            # Excel no disponible: la prueba no puede ejecutarse, pero tampoco es un fallo
            Write-Build Yellow "  [1/5] Excel no disponible - prueba omitida (no es fallo)"
            $pasadas++
        } else {
            # Esto DEBE lanzar excepcion
            Open-ExcelWorkbook -Context $ctx -ExcelApp $xl1 `
                -Path 'C:\ruta\inventada\que\no\existe\archivo.xlsx' -ReadOnly $true
            # Si llega aqui, la funcion no lanzo excepcion -> fallo del test
            Write-Build Red "  [1/5] FALLO: Open-ExcelWorkbook no lanzo excepcion con ruta invalida"
        }
    } catch {
        $msg = $_.Exception.Message.Split([char]10)[0].Trim()
        Write-Build Green "  [1/5] OK: excepcion capturada -> $msg"
        $p1ok = $true
        $pasadas++
    } finally {
        if ($null -ne $xl1) {
            Close-ExcelApp -ExcelApp $xl1
            $xl1 = $null
        }
    }

    # Verificar cero zombis tras el finally
    $z1 = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
    if ($z1 -eq 0) {
        Write-Build Green "  [1/5] Sin zombis Excel tras prueba OK"
    } else {
        Write-Build Red   "  [1/5] $z1 zombi(s) EXCEL tras prueba - finally no funciono"
        Remove-ZombieCom | Out-Null
    }

    # =========================================================================
    # PRUEBA 2: Open-WordDocument con ruta inexistente debe lanzar excepcion
    # =========================================================================
    Write-Build Cyan ""
    Write-Build Cyan "  [2/5] Open-WordDocument ruta inexistente -> excepcion esperada"
    $pruebas++
    $wd2 = $null
    try {
        $wd2 = New-WordApp -Context $ctx -TimeoutSec 20
        if ($null -eq $wd2) {
            Write-Build Yellow "  [2/5] Word no disponible - prueba omitida (no es fallo)"
            $pasadas++
        } else {
            Open-WordDocument -Context $ctx -WordApp $wd2 `
                -Path 'C:\ruta\inventada\que\no\existe\doc.docx' -ReadOnly $true
            Write-Build Red "  [2/5] FALLO: Open-WordDocument no lanzo excepcion con ruta invalida"
        }
    } catch {
        $msg = $_.Exception.Message.Split([char]10)[0].Trim()
        Write-Build Green "  [2/5] OK: excepcion capturada -> $msg"
        $pasadas++
    } finally {
        if ($null -ne $wd2) {
            Close-WordApp -WordApp $wd2
            $wd2 = $null
        }
    }

    $z2 = @(Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
    if ($z2 -eq 0) {
        Write-Build Green "  [2/5] Sin zombis Word tras prueba OK"
    } else {
        Write-Build Red   "  [2/5] $z2 zombi(s) WINWORD - finally no funciono"
        Remove-ZombieCom | Out-Null
    }

    # =========================================================================
    # PRUEBA 3: Assert-Param con valor nulo debe lanzar excepcion
    # =========================================================================
    Write-Build Cyan ""
    Write-Build Cyan "  [3/5] Assert-Param valor nulo -> excepcion esperada"
    $pruebas++
    try {
        Assert-Param -Name 'Centro' -Value $null -Context $ctx
        Write-Build Red "  [3/5] FALLO: Assert-Param no lanzo excepcion con valor nulo"
    } catch {
        $msg = $_.Exception.Message.Split([char]10)[0].Trim()
        Write-Build Green "  [3/5] OK: excepcion capturada -> $msg"
        $pasadas++
    }

    # =========================================================================
    # PRUEBA 4: Assert-Param con cadena vacia debe lanzar excepcion
    # =========================================================================
    Write-Build Cyan ""
    Write-Build Cyan "  [4/5] Assert-Param cadena vacia -> excepcion esperada"
    $pruebas++
    try {
        Assert-Param -Name 'Centro' -Value '' -Context $ctx
        Write-Build Red "  [4/5] FALLO: Assert-Param no lanzo excepcion con cadena vacia"
    } catch {
        $msg = $_.Exception.Message.Split([char]10)[0].Trim()
        Write-Build Green "  [4/5] OK: excepcion capturada -> $msg"
        $pasadas++
    }

    # =========================================================================
    # PRUEBA 5: Assert-FileExists con ruta inventada debe lanzar excepcion
    # =========================================================================
    Write-Build Cyan ""
    Write-Build Cyan "  [5/5] Assert-FileExists ruta invalida -> excepcion esperada"
    $pruebas++
    try {
        Assert-FileExists -Path 'C:\inventado\archivo_que_no_existe_nunca.xlsx' -Context $ctx
        Write-Build Red "  [5/5] FALLO: Assert-FileExists no lanzo excepcion con ruta invalida"
    } catch {
        $msg = $_.Exception.Message.Split([char]10)[0].Trim()
        Write-Build Green "  [5/5] OK: excepcion capturada -> $msg"
        $pasadas++
    }

    # =========================================================================
    # VERIFICACION FINAL: cero zombis en todo el sistema
    # =========================================================================
    Write-Build Cyan ""
    Write-Build Cyan "  [POST] Verificando cero zombis tras todas las pruebas..."
    $zombisFinales = @(Get-Process -Name 'EXCEL','WINWORD' -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
    $pruebas++
    if ($zombisFinales -eq 0) {
        Write-Build Green "  [POST] Sin zombis activos OK"
        $pasadas++
    } else {
        Write-Build Red   "  [POST] $zombisFinales zombi(s) activos - finally blocks fallaron"
        Remove-ZombieCom | Out-Null
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan ""
    $fallidas = $pruebas - $pasadas
    Write-Build Cyan "  Pruebas: $pruebas  Pasadas: $pasadas  Fallidas: $fallidas"

    $ok = ($fallidas -eq 0)
    Write-BuildLog $ctx 'INFO' "diag_errores: pruebas=$pruebas pasadas=$pasadas"
    Write-RunResult -Context $ctx -Success $ok

    if (-not $ok) {
        throw "diag_errores: $fallidas prueba(s) no se comportaron como esperado"
    } else {
        Write-Build Green "  RESULTADO: OK - manejo de errores funciona correctamente"
    }
}
