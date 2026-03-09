# Login SUNAT Menú SOL (Playwright)

Aplicación en Python que abre un navegador controlado con Playwright y realiza el login en la página de SUNAT Menú SOL usando credenciales desde un archivo JSON.

## Requisitos

- Python 3.8+
- Entorno virtual `venv_sunat` (ya creado)

## Configuración

1. **Edita `credenciales.json`** con tus datos reales de SOL:

```json
{
    "ruc": "20123456789",
    "usuario": "TU_USUARIO_SOL",
    "contrasena": "TU_CLAVE_SOL"
}
```

- **ruc**: 11 dígitos del RUC.
- **usuario**: Usuario de SUNAT Operaciones en Línea (SOL).
- **contrasena**: Clave SOL.

2. **No subas `credenciales.json`** a repositorios públicos (contiene datos sensibles).

## Uso

Desde la carpeta del proyecto, con el entorno virtual activado:

**PowerShell:**

```powershell
.\venv_sunat\Scripts\Activate.ps1
python login_sunat.py
```

**CMD:**

```cmd
venv_sunat\Scripts\activate.bat
python login_sunat.py
```

El script abrirá Chromium, cargará la página de login SUNAT, rellenará RUC, usuario y contraseña, y hará clic en "Iniciar sesión". La ventana permanecerá abierta hasta que pulses Enter en la consola.

## Estructura

- `login_sunat.py` – Script principal con Playwright.
- `credenciales.json` – RUC, usuario y contraseña (crear/editar localmente).
- `requirements.txt` – Dependencias (playwright).
- `venv_sunat/` – Entorno virtual con las librerías instaladas.

## Notas

- Si la página de SUNAT cambia el flujo (por ejemplo, primero RUC y luego usuario/contraseña en otra pantalla), puede ser necesario ajustar los selectores o los pasos en `login_sunat.py`.
- Los IDs usados son: `#txtRuc`, `#txtUsuario` (si existe), `#txtContrasena`.
