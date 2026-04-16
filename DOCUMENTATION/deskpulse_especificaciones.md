# DeskPulse Suite — Especificaciones Técnicas

**Versión:** `v1.2.5-four-color-standard-f2`  
**Archivo fuente:** `Deskpulse_suite_complete.py`  
**Plataforma:** Windows (Python 3.10+, Tkinter)  
**Mantenedor:** Evi Support International

---

## Tabla de Contenidos

1. [Resumen General](#1-resumen-general)
2. [Dependencias](#2-dependencias)
3. [Constantes y Configuración Global](#3-constantes-y-configuración-global)
4. [Modelos de Datos](#4-modelos-de-datos)
5. [Capas del Sistema](#5-capas-del-sistema)
   - 5.1 [Capa de Dominio — Helpers y Utilidades](#51-capa-de-dominio--helpers-y-utilidades)
   - 5.2 [Capa de Persistencia](#52-capa-de-persistencia)
   - 5.3 [Capa de Seguridad](#53-capa-de-seguridad)
   - 5.4 [Capa de Monitoreo](#54-capa-de-monitoreo)
   - 5.5 [Capa de Sincronización — Google Sheets](#55-capa-de-sincronización--google-sheets)
   - 5.6 [Capa UI](#56-capa-ui)
6. [Flujo de Navegación](#6-flujo-de-navegación)
7. [Modos de Trabajo](#7-modos-de-trabajo)
8. [Estructura de Archivos Generados](#8-estructura-de-archivos-generados)
9. [Reglas de Negocio Clave](#9-reglas-de-negocio-clave)
10. [Diagrama de Clases Resumido](#10-diagrama-de-clases-resumido)

---

## 1. Resumen General

DeskPulse Suite es una aplicación de escritorio para el monitoreo de actividad de agentes en un entorno BPO. Registra eventos de sesión (clock in/out, lunch, meetings), toma muestras periódicas de actividad (teclado, clics, scroll, app en foco) y exporta los datos en CSV cifrado. Opcionalmente sincroniza eventos a Google Sheets en tiempo real.

La aplicación corre en una sola ventana Tkinter con navegación por vistas. Toda la lógica de muestreo corre en un hilo secundario (`SampleWorker`) para no bloquear la UI.

---

## 2. Dependencias

| Paquete | Uso | Obligatorio |
|---|---|---|
| `tkinter` | UI | Sí (stdlib) |
| `bcrypt` | Hash de contraseñas admin | Recomendado |
| `cryptography` | Cifrado AES-256-GCM de exports | Recomendado |
| `mss` | Capturas de pantalla | Opcional |
| `Pillow` | Procesamiento de imágenes | Opcional |
| `psutil` | Info de procesos (app tracker) | Opcional |
| `pynput` | Listener de teclado y mouse | Opcional |
| `pywin32` | Ventana en primer plano (Windows) | Opcional |
| `gspread` + `google-auth` | Sync Google Sheets | Opcional |

Todas las dependencias opcionales se importan con `try/except` y se deshabilitan graciosamente si no están disponibles (`_CRYPTO_OK`, `_MSS_OK`, `_PYNPUT_OK`, etc.).

---

## 3. Constantes y Configuración Global

### Rutas de Almacenamiento

```
APP_HOME          → %LOCALAPPDATA%/DeskPulseSuite/
├── config.json           # Configuración del agente y admin
├── session_state.json    # Estado de sesión activa (recovery)
├── sheets_queue.json     # Cola de reintentos Google Sheets
├── credentials.json      # Service account GCP
├── records/              # CSV de sesiones por agente/fecha
└── exports/              # ZIPs cifrados de exportación
```

### Paleta de Colores (4-color standard)

| Variable | Hex | Uso |
|---|---|---|
| `C_NAVY` | `#0D2149` | Acento principal, cards oscuras |
| `C_INK` | `#1A1A1A` | Texto principal |
| `C_GRAY` | `#6B7280` | Texto secundario, botones regulares |
| `C_LIGHT` | `#D1D5DB` | Bordes, fondos de botones primarios |
| `C_CHALK` | `#F3F4F6` | Fondo general, texto sobre acento |
| `C_STATUS_ONLINE` | `#16A34A` | Indicador verde ONLINE |
| `C_STATUS_IDLE` | `#DC2626` | Indicador rojo IDLE |

### Tipografía

- **Familia:** `Space Mono`
- Variantes: `F_BASE`, `F_TITLE`, `F_BIG_STATUS`, `F_MUTED`, `F_BOLD`, `F_AGENT_NAME`

### Retry de Google Sheets

- `SHEETS_RETRY_INTERVAL = 60` segundos entre intentos
- `SHEETS_MAX_RETRIES = 10` intentos antes de descartar un evento

### Seguridad Admin

- `ADMIN_MAX_ATTEMPTS = 3` intentos fallidos
- `ADMIN_LOCKOUT_SECONDS = 30` segundos de bloqueo

---

## 4. Modelos de Datos

### `AppConfig` (dataclass)

Representa la configuración cargada desde `config.json`. Todos los campos tienen valores por defecto.

| Campo | Tipo | Descripción |
|---|---|---|
| `first_run_done` | bool | Si ya se completó el setup inicial |
| `agent_id` | str | ID único del agente |
| `agent_name` | str | Nombre completo del agente |
| `project_name` | str | Proyecto/cliente asignado |
| `work_mode` | str | `In Office`, `Home Office` o `Training` |
| `activity_rule` | dict | Umbrales mínimos de actividad por muestra |
| `admin_username` | str | Usuario de admin |
| `admin_password` | str | Hash bcrypt o plain si no está hasheado |
| `admin_password_hashed` | bool | Indica si la contraseña está hasheada |
| `export_password` | str | Contraseña para cifrar exports ZIP |
| `sheets_enabled` | bool | Activar sync Google Sheets |
| `sheets_spreadsheet_link` | str | URL o ID del spreadsheet |
| `sheets_credentials_file` | str | Ruta al `credentials.json` GCP |
| `import_source_file` | str | Ruta del último archivo de config importado |
| `import_imported_at` | str | Timestamp de importación |

### `SessionData` (dataclass)

Representa una sesión de trabajo activa o cerrada.

| Campo | Tipo | Descripción |
|---|---|---|
| `session_id` | str | ID único formato `YYYYMMDDHHMI` |
| `agent_id` | str | ID del agente |
| `agent_name` | str | Nombre del agente |
| `project` | str | Proyecto |
| `work_mode` | str | Modo de trabajo |
| `session_type` | str | `Normal`, `Schedule Change` u `Overtime` |
| `clock_in` | datetime? | Timestamp de entrada |
| `clock_out` | datetime? | Timestamp de salida |
| `lunch_start` | datetime? | Inicio de almuerzo |
| `lunch_end` | datetime? | Fin de almuerzo |
| `lunch_sec` | float | Segundos acumulados de almuerzo |
| `meeting_sec` | float | Segundos acumulados en meetings |
| `status` | str | `IDLE`, `ONLINE`, `LUNCH` o `MEETING` |
| `total_samples` | int | Total de muestras tomadas |
| `active_samples` | int | Muestras que superaron los umbrales |
| `top_app` | str | App más usada en la sesión |
| `app_used` | str | Lista pipe-separated de apps usadas |
| `overtime` | bool | Si se autorizó tiempo extra |
| `overtime_requested_by` | str | Quién autorizó el OT |
| `payable_worked_sec` | float | Segundos pagables (cap de 8h si no hay OT) |
| `overtime_duration_sec` | float | Segundos de OT calculados |

---

## 5. Capas del Sistema

### 5.1 Capa de Dominio — Helpers y Utilidades

Funciones puras de soporte, sin estado.

#### Tiempo y Formato

| Función | Descripción |
|---|---|
| `_now_bolivia()` | Retorna `datetime` actual en UTC-4 (Bolivia) |
| `_fmt(dt)` | Datetime → string `YYYY-MM-DD HH:MM:SS` para CSV/JSON |
| `_fmt_date(dt)` | Datetime → `YYYY-MM-DD` |
| `_fmt_time(dt)` | Datetime → `HH:MM:SS` |
| `_sec_to_hms(seconds)` | Segundos float → string `HH:MM:SS` |
| `_extract_time(value)` | Extrae `HH:MM:SS` de string mixto |
| `_extract_date(value)` | Extrae `YYYY-MM-DD` de string mixto |
| `_format_short_session_id(dt)` | Genera ID de sesión `YYYYMMDDHHMI` |

#### Validación y Normalización

| Función | Descripción |
|---|---|
| `_normalize_work_mode(value)` | Normaliza aliases de modo de trabajo al valor canónico |
| `_normalize_import_key(value)` | Mapea claves de config importada a sus nombres internos |
| `_config_is_ready(cfg)` | Verifica si la configuración está completa para operar |
| `_has_taken_lunch(session)` | Retorna True si el agente ya tomó almuerzo en la sesión |
| `_sanitize_fs_name(value)` | Limpia un string para uso seguro como nombre de carpeta |
| `_agent_folder_label(id, name)` | Compone la etiqueta `{id} - {name}` para la carpeta del agente |
| `_extract_sheet_id(value)` | Extrae el ID de spreadsheet de una URL de Google Sheets |
| `_format_app_list(apps)` | Convierte lista de apps a formato `\|app1\|app2\|` |
| `_should_ignore_tracked_app(proc, exe, title)` | Retorna True si la app debe excluirse del tracker (DeskPulse, sistema) |
| `_build_import_profile(data)` | Valida y construye el perfil de agente desde un dict importado |
| `load_import_settings_file(path)` | Carga config de agente desde archivo JSON o key=value |

#### Assets y Entorno Windows

| Función | Descripción |
|---|---|
| `_runtime_asset_dirs()` | Retorna lista de directorios donde buscar assets (source, PyInstaller) |
| `_resolve_runtime_asset(*names)` | Resuelve la primera ruta existente de una lista de nombres de asset |
| `_enable_dpi_awareness()` | Activa Per-Monitor DPI en Windows |
| `_apply_windows_app_id()` | Setea AppUserModelID para ícono correcto en taskbar |
| `_apply_window_icon(win)` | Aplica `app.ico` a cualquier ventana Tkinter |
| `_center_window(win, w, h)` | Centra una ventana en la pantalla |
| `_apply_theme(root)` | Aplica el tema visual global a todos los widgets ttk |

---

### 5.2 Capa de Persistencia

#### `ConfigManager`

Carga, hidrata y guarda `config.json`. Es el punto único de acceso a `AppConfig`.

| Método | Descripción |
|---|---|
| `load()` | Carga config desde disco, crea defaults si no existe |
| `_hydrate()` | Mapea el JSON raw a los campos de `AppConfig` |
| `save_agent(id, name, project)` | Guarda datos del agente y marca `first_run_done` |
| `save_work_mode(mode)` | Actualiza el modo de trabajo |
| `save_imported_profile(profile, source)` | Guarda perfil importado con metadatos de importación |
| `save_credentials(user, pw_hash, exp_pw, hashed)` | Actualiza credenciales admin y export password |
| `save_google_sheets(enabled, link, cred_file)` | Guarda configuración de Google Sheets |

#### `StateManager`

Persiste y recupera el estado de la sesión activa en `session_state.json`. Permite recovery si la app se cierra inesperadamente.

| Método | Descripción |
|---|---|
| `save(session)` | Serializa `SessionData` a JSON |
| `load()` | Lee el JSON y retorna el dict, o `None` si no existe |
| `delete()` | Elimina el archivo de estado al cerrar sesión normalmente |

#### `SessionLogger`

Escribe los archivos CSV de registro. Todos los métodos son estáticos.

| Método | Descripción |
|---|---|
| `_session_dir(session)` | Retorna ruta `records/{agent}/{year}/{month}/{day}/{session_id}/` |
| `_screenshots_dir(session)` | Retorna subcarpeta `screenshots/` dentro del directorio de sesión |
| `_ensure_dir(session)` | Crea el árbol de carpetas de la sesión |
| `_ensure_csv_schema(path, fieldnames)` | Migra un CSV existente si el schema cambió |
| `append_sample(session, row)` | Agrega una fila a `activity_samples.csv` |
| `write_session_log(session, ...)` | Escribe la fila de cierre en `session_log.csv` |
| `write_summary_day(session, ...)` | Agrega fila al `summary_day.csv` del día |

**Columnas de `activity_samples.csv`:**
`session_id`, `agent_name`, `sample_date`, `sample_time`, `status`, `keystrokes`, `clicks`, `scroll_events`, `app_used`, `connection_type`, `network_name`, `activity_flag`, `screenshot_taken`

**Columnas de `session_log.csv`:**
`session_id`, `date`, `agent_id`, `agent_name`, `project`, `work_mode`, `session_type`, `clock_in`, `clock_out`, `lunch_start`, `lunch_end`, `lunch_time`, `meeting_time`, `net_worked_time`, `total_time_worked`, `payable_time_worked`, `overtime_duration`, `overtime_requested_by`, `overtime`, `ot_note`, `top_app`, `app_used`, `activity_avg`, `close_reason`

---

### 5.3 Capa de Seguridad

#### `SecurityUtils`

| Método | Descripción |
|---|---|
| `hash_password(plain)` | Genera hash bcrypt con salt aleatorio |
| `verify_password(plain, stored, is_hashed)` | Verifica contraseña contra hash bcrypt o texto plano |
| `encrypt_file(src, dest, password)` | Cifra un archivo con AES-256-GCM + PBKDF2-SHA256 (260,000 iteraciones). Formato: `salt(16) + nonce(12) + ciphertext+tag` |

---

### 5.4 Capa de Monitoreo

#### `InputMonitor`

Captura eventos de teclado y mouse con `pynput`. Todos los contadores son thread-safe y se resetean en cada `snapshot()`.

| Método | Descripción |
|---|---|
| `start()` | Inicia los listeners de teclado y mouse como daemons |
| `stop()` | Detiene ambos listeners |
| `snapshot()` | Retorna dict `{keystrokes, clicks, scroll_events, mouse_move_pct}` y resetea contadores |

Internamente usa una grilla de 20×20 celdas para calcular la cobertura de movimiento del mouse (`mouse_move_pct`).

#### `AppTracker`

Detecta la aplicación en primer plano usando `win32gui` + `psutil`. Mantiene la última app válida para casos donde no se puede obtener el foco.

| Método | Descripción |
|---|---|
| `get_visible_apps()` | Retorna lista con la app en foco actual, excluyendo sistema y DeskPulse |
| `_fallback()` | Retorna la última app válida registrada si no se puede obtener el foco actual |

Excluye automáticamente: procesos de sistema Windows, tool windows, ventanas propias de DeskPulse, y procesos con rutas en `C:\Windows\System32` y similares.

#### `SampleWorker`

Hilo de fondo que ejecuta el ciclo de muestreo periódico.

| Método | Descripción |
|---|---|
| `start()` | Inicia el hilo y el `InputMonitor` |
| `stop()` | Detiene el hilo y el `InputMonitor` |
| `_run()` | Loop principal: duerme hasta el próximo intervalo y llama `_take_sample()` |
| `_take_sample()` | Recolecta métricas, evalúa actividad, guarda CSV, persiste estado |
| `set_status(new_status, now)` | Cambia el estado de la sesión validando transiciones permitidas |

**Transiciones de estado válidas:**

```
IDLE    → ONLINE
ONLINE  → LUNCH, MEETING
LUNCH   → ONLINE
MEETING → ONLINE
```

**Lógica de `activity_flag`:**
Una muestra se marca `ACTIVE` si cumple los umbrales configurados en `activity_rule`:
- `min_keystrokes` (default: 20)
- `min_clicks` (default: 30)
- `min_scroll` (default: 1)
- `min_mouse_pct` (default: 20%)

---

### 5.5 Capa de Sincronización — Google Sheets

#### `SheetsSync`

Envía eventos de sesión a una hoja de Google Sheets. Si el envío falla, encola el evento en `sheets_queue.json` y lo reintenta cada 60 segundos en un hilo daemon.

| Método | Descripción |
|---|---|
| `refresh_config(cfg)` | Actualiza la config y fuerza reconexión al cliente gspread |
| `send_event(event_type, session)` | Envía o encola un evento de sesión |
| `_get_client()` | Inicializa (o reutiliza) el cliente `gspread` con service account |
| `_worksheet(gc)` | Obtiene o crea la hoja con el nombre del agente; verifica/actualiza headers |
| `_push(entry)` | Escribe o actualiza una fila en Sheets; retorna True si exitoso |
| `_retry_loop()` | Loop daemon que reintenta eventos encolados, descarta los que superan `SHEETS_MAX_RETRIES` |

**Columnas de la hoja:** `session_id`, `date`, `agent_id`, `agent_name`, `session_type`, `clock_in`, `clock_out`, `lunch_start`, `lunch_end`, `status`, `net_worked_time`

---

### 5.6 Capa UI

#### `App` (ventana principal)

Subclase de `tk.Tk`. Orquesta la navegación entre vistas, el `ConfigManager` y el `SheetsSync`.

| Método | Descripción |
|---|---|
| `_show_splash()` | Lanza el `SplashScreen` antes de mostrar la ventana principal |
| `_navigate(target)` | Destruye la vista actual y construye la vista de destino |
| `_on_close()` | Muestra confirmación de cierre; fuerza clock-out si hay sesión activa |

**Vistas registradas:** `start`, `admin_login`, `admin_console`, `agent`

#### `SplashScreen`

Ventana sin bordes que se muestra 3 segundos al inicio con fade-out de 600ms. Muestra el logo si `splash.png/jpg/gif` existe, o fallback de texto.

#### `StartView`

Vista de selección de rol. Verifica `_config_is_ready()` antes de permitir acceso a la vista de agente.

#### `AdminLoginView`

Formulario de login con lockout: bloquea el acceso 30 segundos tras 3 intentos fallidos. Verifica la contraseña con `SecurityUtils.verify_password()`.

#### `AdminConsoleView`

Panel de administración con tres pestañas:

| Pestaña | Función |
|---|---|
| **Settings** | Importar archivo de config de agente (JSON o key=value). Muestra preview de la config activa. |
| **Credentials** | Cambiar usuario/contraseña admin (bcrypt), contraseña de export, y configurar Google Sheets |
| **Export** | (Ver `ExportDialog`) |

#### `AgentView`

Vista principal del agente. Muestra estado actual, timers en vivo y botones de acción.

| Método | Descripción |
|---|---|
| `_clock_in()` | Inicia sesión: crea `SessionData`, arranca `SampleWorker`, inicia tick de UI |
| `_clock_out(force)` | Cierra la sesión: calcula tiempos, aplica cap OT, escribe CSVs, envía a Sheets |
| `_toggle_lunch()` | Alterna estado LUNCH/ONLINE via `SampleWorker.set_status()` |
| `_toggle_meeting()` | Alterna estado MEETING/ONLINE |
| `_update_totals()` | Actualiza labels de tiempo online/lunch/meeting cada segundo |
| `_refresh_status_visuals()` | Actualiza colores e indicadores según el estado actual |
| `_refresh_action_buttons()` | Habilita/deshabilita botones según estado y restricciones de negocio |
| `_schedule_status_pulse()` | Efecto de pulso visual en 850ms para estados MEETING |
| `force_clockout_if_active()` | Clock-out forzado al cerrar con X; llamado por `App._on_close()` |

#### `ExportDialog`

Diálogo de exportación de registros. Permite seleccionar rango de fechas y agente. Genera un ZIP con los CSV, cifrado con AES-256-GCM usando la export password configurada.

---

## 6. Flujo de Navegación

```
App init
  └─► SplashScreen (3s + fade)
        └─► StartView
              ├─► [Agente] → _config_is_ready? → AgentView
              │                                      ├─ Clock In → SampleWorker loop
              │                                      ├─ Lunch / Meeting toggle
              │                                      └─ Clock Out → CSVs + Sheets
              └─► [Admin] → AdminLoginView (lockout 3 intentos)
                              └─► AdminConsoleView
                                    ├─ Tab Settings: import config
                                    ├─ Tab Credentials: cambiar contraseñas + Sheets
                                    └─ Tab Export: exportar CSVs cifrados
```

---

## 7. Modos de Trabajo

| Modo | Intervalo muestreo | Screenshots |
|---|---|---|
| `In Office` | 5 min | Obligatorias en clock in/out |
| `Home Office` | 2 min | En inactividad detectada + clock in/out |
| `Training` | 10 min | Obligatorias en clock in/out |

El umbral de inactividad para Home Office es `HOME_OFFICE_IDLE_THRESHOLD_SEC = 120` segundos.

---

## 8. Estructura de Archivos Generados

```
%LOCALAPPDATA%/DeskPulseSuite/
├── config.json
├── session_state.json          ← solo existe mientras hay sesión activa
├── sheets_queue.json
├── credentials.json
├── records/
│   └── {agent_id} - {agent_name}/
│       └── {YYYY}/
│           └── {MM}/
│               └── {DD}/
│                   ├── summary_day.csv
│                   └── {session_id}/
│                       ├── session_log.csv
│                       ├── activity_samples.csv
│                       └── screenshots/
│                           └── *.png
└── exports/
    └── export_{timestamp}.zip.enc   ← AES-256-GCM cifrado
```

---

## 9. Reglas de Negocio Clave

### Cap de Horas Pagables (OT Logic)

- Jornada regular: `REGULAR_WORKDAY_SEC = 8 horas`
- Si la sesión supera las 8h y **no se autorizó OT**, el tiempo pagable se capea a 8h
- El OT debe ser autorizado por roles en `OVERTIME_REQUESTERS`: `CLIENT`, `TEAM LEADER`, `OFFICE OPERATIONS`
- `overtime_duration` = tiempo trabajado más allá de las 8h (solo si fue autorizado)

### Midnight-Safe Sessions

Las sesiones persisten en `session_state.json` y sobreviven reinicios. Al relanzar la app con un estado guardado, se recupera la sesión donde quedó.

### Self-Ignore del App Tracker

DeskPulse excluye su propio proceso del rastreo de aplicaciones mediante `APP_TRACKER_IGNORE_NAME_KEYWORDS` y `APP_TRACKER_IGNORE_TITLE_KEYWORDS`, ambos apuntando a la keyword `"deskpulse"`.

### Migración de Schema CSV

`SessionLogger._ensure_csv_schema()` detecta si el schema de un CSV existente no coincide con el esperado, y migra las filas preservando los datos con columnas renombradas o nuevas.

### Import de Config de Agente

`load_import_settings_file()` acepta dos formatos:
- **JSON:** `{"agent_id": "...", "agent_name": "...", ...}`
- **Key-value:** `Agent ID=1234` o `agent_name: Nombre`

Soporta múltiples aliases para cada campo (ver `CONFIG_IMPORT_KEY_ALIASES` y `WORK_MODE_ALIASES`).

---

## 10. Diagrama de Clases Resumido

```
App (tk.Tk)
 ├── ConfigManager ──► AppConfig (dataclass)
 ├── SheetsSync
 └── Views
      ├── StartView
      ├── AdminLoginView
      ├── AdminConsoleView
      └── AgentView
           └── SampleWorker (Thread)
                ├── InputMonitor (pynput)
                ├── AppTracker (win32gui + psutil)
                └── SessionLogger (static) ──► CSVs en disco

SecurityUtils (static)
  ├── hash_password / verify_password (bcrypt)
  └── encrypt_file (AES-256-GCM)

StateManager (static)
  └── session_state.json (crash recovery)
```

---

*Documento generado automáticamente desde el análisis de `Deskpulse_suite_complete.py` — DeskPulse Suite v1.2.5*
