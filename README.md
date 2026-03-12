# AutoBuild v1.0
## Motor de Automatizacion Corporativa - PowerShell 5.1 + Invoke-Build

Motor modular de automatizacion para entornos de produccion restringidos.
Portable: no requiere instalacion de modulos externos en el sistema.

**Version:** 1.0
**Motor:** Invoke-Build 5.14.23 (incluido en tools/InvokeBuild/)
**Requiere:** PowerShell 5.1 (Windows PowerShell, NO PowerShell Core 7)
**SO:** Windows 10/11, Windows Server 2016+
**Entorno:** Office 16, SAP GUI 800 (8000.1.11.1161)

---

## INDICE

1. [Inicio rapido](#1-inicio-rapido)
2. [Estructura del proyecto](#2-estructura-del-proyecto)
3. [Uso desde consola](#3-uso-desde-consola)
4. [Crear una nueva tarea](#4-crear-una-nueva-tarea)
5. [Librerias disponibles](#5-librerias-disponibles)
6. [Configuracion](#6-configuracion)
7. [Tareas con checkpoint](#7-tareas-con-checkpoint)
8. [Bugs conocidos y reglas obligatorias](#8-bugs-conocidos-y-reglas-obligatorias)
9. [Troubleshooting](#9-troubleshooting)
10. [Interfaz grafica (UI)](#10-interfaz-grafica-ui)

---

## 1. INICIO RAPIDO

```powershell
# 1. Navegar al proyecto
cd C:\ruta\a\AutoBuild

# 2. Setup inicial (una sola vez por maquina)
.\Setup.ps1

# 3. Ver tareas disponibles
.\Run.ps1 -List

# 4. Ejecutar diagnostico del entorno
.\Run.ps1 diagnostico

# 5. Ejecutar una tarea con parametros
.\Run.ps1 sap_stock -Centro 1000
.\Run.ps1 sap_stock -Centro 1000 -Almacen 0001

# 6. Generar reporte Excel desde CSV
#    Copiar el CSV a input/ primero
.\Run.ps1 excel_reporte -Extra datos.csv

# 7. Limpiar procesos COM zombi
.\Run.ps1 limpiar_com
```

IMPORTANTE antes de tareas COM (Excel/Word/SAP):
Abrir la aplicacion manualmente, cerrar cualquier dialogo pendiente
(recuperacion de archivos, actualizaciones, asistente de activacion)
y cerrar limpias.

---

## 2. ESTRUCTURA DEL PROYECTO

```
AutoBuild/
|-- Run.ps1                  <- Punto de entrada principal
|-- Setup.ps1                <- Configuracion inicial (ejecutar una vez)
|-- New-Task.ps1             <- Generador de nuevas tareas
|-- engine.config.json       <- Configuracion global
|
|-- engine/
|   `-- Main.build.ps1       <- Orquestador Invoke-Build (carga tareas automaticamente)
|
|-- lib/
|   |-- Logger.ps1           <- Logging JSONL, carga de config
|   |-- Context.ps1          <- Fabrica de contextos de ejecucion
|   |-- ComHelper.ps1        <- COM seguro con timeout (BUG-COM-FREEZE-01)
|   |-- SapHelper.ps1        <- SAP GUI Scripting
|   |-- ExcelHelper.ps1      <- Excel COM con liberacion correcta
|   `-- Assertions.ps1       <- Validaciones y prerequisitos
|
|-- tasks/
|   |-- task_TEMPLATE.ps1    <- Plantilla para nuevas tareas
|   |-- task_sap_stock.ps1   <- Ejemplo: stock SAP (MMBE)
|   `-- task_excel_reporte.ps1 <- Ejemplo: CSV a Excel
|
|-- tools/
|   |-- InvokeBuild/         <- Invoke-Build 5.14.23 portable
|   |   |-- Invoke-Build.ps1
|   |   |-- Build-Checkpoint.ps1
|   |   `-- Build-Parallel.ps1
|   `-- Test-Ascii.ps1       <- Verificador de ASCII (BUG-ASCII-01)
|
|-- input/                   <- Archivos de entrada (CSV, etc.)
|-- output/                  <- Archivos de salida intermedios
|-- reports/                 <- Reportes generados
`-- logs/
    |-- registry.jsonl       <- Log estructurado de todas las ejecuciones
    `-- checkpoint_*.clixml  <- Checkpoints de builds persistentes
```

---

## 3. USO DESDE CONSOLA

```powershell
# Ver tareas disponibles con sinopsis
.\Run.ps1 -List

# Ejecutar tarea
.\Run.ps1 NOMBRE_TAREA

# Ejecutar tarea con parametros
.\Run.ps1 NOMBRE_TAREA -Centro 1000 -Almacen 0001

# Preview de lo que haria sin ejecutar
.\Run.ps1 NOMBRE_TAREA -WhatIf

# Ejecutar con checkpoint (reanudar si falla)
.\Run.ps1 NOMBRE_TAREA -Checkpoint

# Reanudar tarea interrumpida
.\Run.ps1 NOMBRE_TAREA -Resume
```

### Parametros disponibles en Run.ps1

| Parametro    | Tipo   | Descripcion                          |
|--------------|--------|--------------------------------------|
| -Centro      | string | Centro SAP (ej: 1000)                |
| -Almacen     | string | Almacen SAP (ej: 0001)               |
| -Fecha       | string | Fecha en formato que necesite la tarea|
| -Extra       | string | Parametro de proposito general       |
| -Checkpoint  | switch | Ejecutar con checkpoint activado      |
| -Resume      | switch | Reanudar desde checkpoint guardado    |
| -WhatIf      | switch | Preview sin ejecutar                  |
| -List        | switch | Mostrar tareas disponibles            |

---

## 4. CREAR UNA NUEVA TAREA

### Opcion A - Automatico (recomendado)

```powershell
.\New-Task.ps1 -Name sap_ventas -Category SAP -Description "Ventas por periodo"
# Genera: tasks\task_sap_ventas.ps1
# Editar e implementar la logica dentro del bloque task sap_ventas { }
```

### Reglas obligatorias

- Solo ASCII en el archivo .ps1 (BUG-ASCII-01)
- Nombre del archivo: task_[a-zA-Z0-9_-]+.ps1
- La tarea se registra con: task NOMBRE { ... }
- Siempre crear contexto con New-TaskContext al inicio
- No hardcodear rutas: usar $ctx.Paths.*
- No definir funciones globales: poner en lib/
- Tareas COM: New-ExcelApp/Invoke-ComWithTimeout + try/finally
- Llamar Write-RunResult al terminar correctamente

### Estructura minima

```powershell
# Synopsis: Lo que aparece en .\Run.ps1 -List
task mi_tarea {
    $ctx = New-TaskContext `
        -TaskName 'mi_tarea' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Inicio'

    # ... logica ...

    Write-RunResult -Context $ctx -Success $true
}
```

---

## 5. LIBRERIAS DISPONIBLES

### Logger.ps1
- `Get-EngineConfig -Root`          Lee engine.config.json y crea carpetas
- `Write-BuildLog $ctx LEVEL msg`   Escribe en consola y en logs/registry.jsonl
- `Write-RunResult $ctx $success`   Registra resultado final de la tarea

### Context.ps1
- `New-TaskContext -TaskName -Config -Root [-Params]`   Crea contexto de ejecucion

### ComHelper.ps1
- `Test-ComAvailable -ProgId [-TimeoutSec]`              Comprueba si COM responde
- `Invoke-ComWithTimeout $ctx -ProgId [-TimeoutSec]`     Instancia COM con seguridad
- `Release-ComObject $obj`                               Libera un objeto COM
- `Invoke-ComCleanup -Document -Application $ctx`        Limpieza completa COM
- `Remove-ZombieCom`                                     Mata procesos COM sin ventana

### SapHelper.ps1
- `Get-SapSession $ctx`                                  Obtiene sesion SAP activa
- `Invoke-SapTransaction $ctx $sess -TCode`              Ejecuta transaccion
- `Set-SapField $ctx $sess -FieldId -Value`              Escribe en campo SAP
- `Get-SapField $ctx $sess -FieldId`                     Lee campo SAP
- `Invoke-SapButton $ctx $sess -ButtonId`                Presiona boton SAP
- `Export-SapTableToArray $ctx $sess -TableId -Columns`  Exporta tabla SAP

### ExcelHelper.ps1
- `New-ExcelApp $ctx [-TimeoutSec]`                      Crea instancia Excel segura
- `New-ExcelWorkbook $ctx $xl`                           Crea workbook vacio
- `Open-ExcelWorkbook $ctx $xl -Path [-ReadOnly]`        Abre workbook existente
- `Save-ExcelWorkbook $ctx $wb -Path [-Format]`          Guarda workbook
- `Write-ExcelData $ctx $ws -Data [-StartRow] [-Columns]` Escribe datos en hoja
- `Close-ExcelWorkbook $wb [-Save]`                      Cierra y libera workbook
- `Close-ExcelApp $xl`                                   Cierra y libera Excel

### Assertions.ps1
- `Assert-Param -Name -Value $ctx`                       Valida parametro no vacio
- `Assert-FileExists -Path $ctx`                         Valida que existe archivo
- `Assert-SapSession $session $ctx`                      Valida sesion SAP
- `Test-TaskAsset $ctx -Files @{} -Params @{}`           Valida multiples prerequisitos

---

## 6. CONFIGURACION

```json
{
  "engine":  { "logLevel": "INFO", "maxRetries": 3, "retryDelaySeconds": 5 },
  "sap":     { "systemId": "PRD", "client": "800", "language": "ES", "timeout": 180 },
  "excel":   { "visible": false, "screenUpdating": false },
  "reports": { "defaultFormat": "xlsx", "retentionDays": 30 }
}
```

logLevel puede ser: DEBUG | INFO | WARN | ERROR

NO guardar credenciales en JSON. Usar Windows Credential Manager:

```powershell
# Guardar una vez por usuario (no queda en logs)
cmdkey /generic:AutoBuild_SAP /user:TU_USUARIO /pass:TU_PASSWORD

# Leer en la tarea
$cred = [System.Net.NetworkCredential]::new('',
    (cmdkey /list:AutoBuild_SAP | Select-String "Usuario")).Password
```

---

## 7. TAREAS CON CHECKPOINT

Para tareas largas que pueden interrumpirse y reanudarse:

```powershell
# Iniciar con checkpoint
.\Run.ps1 sap_stock -Centro 1000 -Checkpoint

# Si falla o se interrumpe, reanudar desde donde quedo
.\Run.ps1 sap_stock -Resume
```

El checkpoint se guarda en logs/checkpoint_TAREA.clixml.
Se elimina automaticamente si la tarea termina con exito.

---

## 8. BUGS CONOCIDOS Y REGLAS OBLIGATORIAS

### BUG-ASCII-01 | Caracteres no-ASCII rompen PS 5.1

PS 5.1 puede fallar con ParserError si hay acentos, em-dash, o cualquier
caracter fuera de ASCII en un .ps1.

**Verificar:**
```powershell
.\tools\Test-Ascii.ps1
# Con correccion automatica de caracteres comunes:
.\tools\Test-Ascii.ps1 -Fix
```

**Regla:** Solo ASCII en todos los archivos .ps1.

### BUG-COM-FREEZE-01 | Excel/Word se congela al instanciar

`New-Object -ComObject Excel.Application` puede bloquearse si Excel tiene
un dialogo modal pendiente (recuperacion de archivos, actualizacion, activacion).

**Solucion:** Usar siempre `New-ExcelApp` (incluye timeout de 30s).
Antes de la primera tarea del dia: abrir Excel, cerrar dialogos, cerrar.

### Reglas COM obligatorias

1. Siempre usar New-ExcelApp o Invoke-ComWithTimeout (nunca New-Object directo)
2. Siempre usar try/finally en bloques COM
3. Orden de liberacion: documento -> aplicacion -> GC
4. Cada llamada del finally envuelta en try/catch individual
5. Configurar modo silencioso antes de operar (lo hace New-ExcelApp)
6. No abrir mas de una instancia COM a la vez

---

## 9. TROUBLESHOOTING

| Sintoma | Causa probable | Solucion |
|---------|---------------|----------|
| ParserError al arrancar | Caracteres no-ASCII en .ps1 | .\tools\Test-Ascii.ps1 -Fix |
| Tarea congelada >5 min | Dialogo modal de Excel/Word | Matar EXCEL.EXE desde Administrador |
| Excel no disponible | Dialogo pendiente o no instalado | Abrir Excel, cerrar dialogos, cerrar |
| SAP no disponible | SAP GUI no iniciado | Iniciar SAP GUI y conectarse |
| Checkpoint no reanuda | Build script modificado | Borrar el .clixml y reiniciar |
| Archivo no encontrado | Falta en input/ | Copiar el archivo a input/ |
| Tarea no aparece en -List | Error en el .ps1 | Abrir en ISE y ejecutar para ver error |
| Procesos zombi | COM no liberado correctamente | .\Run.ps1 limpiar_com |

---

## 10. INTERFAZ GRAFICA (UI)

La interfaz WPF permite operar AutoBuild sin necesidad de consola.

### Requisitos

| Requisito      | Minimo          | Notas                              |
|----------------|-----------------|-------------------------------------|
| PowerShell     | 5.1 Desktop     | Core edition tiene COM limitado     |
| Windows        | 10 / 11         | .NET Framework 4.x requerido        |
| .NET Framework | 4.6+            | Para los ensamblados WPF            |
| AutoBuild      | v1.0+           | Run.ps1 debe estar en el mismo dir  |

### Lanzar la interfaz

```powershell
# Opcion 1 - doble clic en el explorador
Launch-AutoBuildUI.bat

# Opcion 2 - PowerShell (recomendado)
.\Start-AutoBuildUI.ps1                            # rol Operator (defecto)
.\Start-AutoBuildUI.ps1 -Role Developer
.\Start-AutoBuildUI.ps1 -Role Admin
.\Start-AutoBuildUI.ps1 -Role Admin -EnginePath "C:\AutoBuild"

# Opcion 3 - directo (requiere STA manual)
powershell.exe -STA -File AutoBuild.UI.ps1 -Role Admin
```

> Start-AutoBuildUI.ps1 detecta si el hilo es STA y relanza automaticamente si no lo es.

### Control de acceso (RBAC)

| Funcion                         | Operator | Developer | Admin |
|---------------------------------|:--------:|:---------:|:-----:|
| Catalogo de tareas              |    SI    |    SI     |  SI   |
| Ejecutar tareas                 |    SI    |    SI     |  SI   |
| Ver historial y artefactos      |    SI    |    SI     |  SI   |
| Ver metricas y diagnosticos     |    SI    |    SI     |  SI   |
| Crear / editar tareas           |    NO    |    SI     |  SI   |
| Editar engine.config.json       |    NO    |    NO     |  SI   |
| Eliminar artefactos             |    NO    |    NO     |  SI   |
| Gestionar checkpoints           |    NO    |    NO     |  SI   |
| Ver log de auditoria            |    NO    |    NO     |  SI   |

### Paginas disponibles

| Pagina                  | Descripcion                                                   |
|-------------------------|---------------------------------------------------------------|
| Task Catalog            | Lista tasks con filtro por nombre y categoria                 |
| Execute Task            | Formulario de parametros dinamico + salida en tiempo real     |
| Live Monitor            | Trabajos activos + tail del registry.jsonl (refresco 5s)     |
| Execution History       | Historial agrupado por RunId, exportable a CSV                |
| Checkpoint Manager      | Lista y gestiona archivos logs/checkpoint_*.clixml            |
| Artifact Repository     | Navega output/ y reports/, abre, descarga o elimina           |
| Metrics & Observability | KPIs: total, exito, duracion media, tarea mas frecuente       |
| Environment Diagnostics | Valida PS, motor, carpetas, COM (Excel/Word/SAP), herramientas|
| Configuration           | Editor JSON de engine.config.json con validacion              |
| Create New Task         | Llama New-Task.ps1 con preview del archivo generado           |
| Audit Log               | Log de auditoria logs/audit.jsonl, exportable a CSV           |

### Archivos del componente UI

| Archivo                       | Proposito                                                     |
|-------------------------------|---------------------------------------------------------------|
| `AutoBuild.UI.ps1`            | Aplicacion WPF principal (toda la logica, RBAC, datos)        |
| `Start-AutoBuildUI.ps1`       | Lanzador inteligente (modo STA, validaciones)                 |
| `Launch-AutoBuildUI.bat`      | Lanzador Windows para doble clic                              |
| `Invoke-UITask.ps1`           | Wrapper delgado alrededor de Run.ps1 para ejecuciones UI      |
| `Invoke-RetentionCleanup.ps1` | Politica de retencion de artefactos                           |

### Retencion de artefactos

```powershell
# Preview (sin eliminar)
.\Invoke-RetentionCleanup.ps1 -EnginePath "C:\AutoBuild" -WhatIf

# Ejecutar limpieza
.\Invoke-RetentionCleanup.ps1 -EnginePath "C:\AutoBuild"
```

Periodo configurado en engine.config.json (`reports.retentionDays`). Defecto: 30 dias.

### Troubleshooting UI

| Sintoma                            | Causa                        | Solucion                                       |
|------------------------------------|------------------------------|------------------------------------------------|
| Ventana no aparece / crash inicial | Modo MTA en lugar de STA     | Usar Start-AutoBuildUI.ps1 (gestiona STA)      |
| Botones de nav no responden        | Error de scope en el handler | Actualizar a la version corregida del UI       |
| "WPF assemblies not available"     | .NET Framework no instalado  | Instalar .NET Framework 4.6+                   |
| "Run.ps1 not found"                | UI fuera del directorio raiz | Copiar UI al mismo directorio que Run.ps1      |
| Pagina Diagnostics no carga        | COM check tarda >30s         | Normal si Excel/Word no estan instalados       |
| Error de politica de ejecucion     | Execution policy restrictiva | Set-ExecutionPolicy -Scope CurrentUser RemoteSigned |
