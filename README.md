# AutoBuild v3.3 Deep Audit

**AutoBuild v3.3 Deep Audit** es un motor de automatización corporativa avanzado para entornos de producción, diseñado para diagnósticos exhaustivos, auditoría de sistemas y remediación automatizada. Este módulo se centra en análisis profundo de la infraestructura AutoBuild y corrección de fallas críticas, incluyendo integridad del motor, compatibilidad COM, y control de procesos Excel.

---

## Índice

- [Características](#características)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Estructura del Proyecto](#estructura-del-proyecto)
- [Uso](#uso)
- [Tareas de Diagnóstico](#tareas-de-diagnóstico)
- [Remediación Automática](#remediación-automática)
- [Registro de Errores y Logs](#registro-de-errores-y-logs)
- [Contribuciones](#contribuciones)
- [Licencia](#licencia)

---

## Características

- Diagnóstico profundo del motor AutoBuild, incluyendo verificación de módulos y configuraciones.
- Auditoría completa de procesos Excel (COM), control de PID, y prevención de procesos zombie.
- Comprobación de integridad de archivos y rutas críticas de ejecución.
- Generación de reportes CSV y Excel con resultados de auditoría.
- Remediación automatizada de errores comunes en entornos de producción.
- Soporte multiplataforma dentro de entornos Windows PowerShell 5.1+ sin dependencias externas.

---

## Requisitos

- **Sistema Operativo:** Windows 10/11 o Windows Server 2016+
- **PowerShell:** 5.1 o superior
- **Permisos:** Acceso de escritura en directorios de ejecución y configuración
- **Dependencias incluidas:**  
  - `Invoke-Build 5.14.23` (incluido en `tools/InvokeBuild/`)  
  - Módulos internos de AutoBuild v3.3

---

## Instalación

1. Clonar el repositorio:

```powershell
git clone https://github.com/Dan-Castello/AutoBuild_v3_3_deep_audit.git
cd AutoBuild_v3_3_deep_audit

Configurar los permisos de ejecución de scripts:

Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

Ejecutar el diagnóstico inicial:

.\Run.ps1 diag_full_audit
Estructura del Proyecto
AutoBuild_v3_3_deep_audit/
│
├─ engine/                 # Núcleo de funciones y motor de AutoBuild
│  ├─ Main.build.ps1
│  ├─ Config.ps1
│  └─ HelperFunctions.ps1
│
├─ tasks/                  # Tareas de diagnóstico y auditoría
│  ├─ task_diag_engine.ps1
│  ├─ task_diag_excel.ps1
│  ├─ task_diag_csv.ps1
│  └─ task_deep_audit.ps1
│
├─ remediation/            # Scripts de corrección automática
│  ├─ Config.ps1
│  └─ task_diag_engine.ps1
│
├─ ui/                     # Interfaz opcional de ejecución
│  └─ AutoBuild.UI.ps1
│
├─ tools/                  # Herramientas externas incluidas
│  └─ InvokeBuild/
│
└─ README.md
Uso

Ejecutar cualquier tarea desde CLI:

.\Run.ps1 <task_name>

Tareas disponibles:

Task Name	Descripción
diag_csv	Diagnóstico de CSV: lectura/escritura, codificación, delimitadores, 10k filas
diag_engine	Diagnóstico del motor: logging, retry, mutex, filesystem
diag_excel	Diagnóstico Excel COM: disponibilidad, velocidad de escritura, PID tracking
diag_full_audit	Auditoría profunda combinando todas las tareas anteriores
Tareas de Diagnóstico

diag_engine: Verifica la integridad del motor, rutas, logging, y estado de mutex.

diag_excel: Evalúa la capacidad de crear y manipular archivos Excel vía COM. Detecta procesos zombie.

diag_csv: Comprueba lectura y escritura de CSV, codificación, delimitadores y caracteres especiales.

diag_full_audit: Ejecuta un diagnóstico completo, consolidando resultados en un reporte.

Remediación Automática

Scripts dentro de remediation/ permiten corregir:

Configuraciones faltantes o corruptas.

Errores de inicialización de Excel COM.

Archivos críticos ausentes.

Se recomienda ejecutar después del diagnóstico para aplicar correcciones.

Copy-Item remediation\Config.ps1 engine\Config.ps1 -Force
Copy-Item remediation\task_diag_engine.ps1 tasks\task_diag_engine.ps1 -Force
Registro de Errores y Logs

Logs de diagnóstico generados en logs/ (crear si no existe).

Formato CSV para resultados tabulares.

Detalle de errores de ejecución de tareas, incluyendo PID y objetos COM.

Contribuciones

Se aceptan correcciones de errores, mejoras en remediación y nuevas tareas de auditoría.

Mantener compatibilidad con PowerShell 5.1+.

Seguir convenciones de nombres y estructura de carpetas.

Licencia

AutoBuild v3.3 Deep Audit está bajo licencia MIT. Puede ser usado, modificado y redistribuido con atribución al autor original.
