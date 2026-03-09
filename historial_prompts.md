# Desarrollo de aplicación con Playwright
_Exported on 3/8/2026 at 22:54:34 GMT-5 from Cursor (2.6.14)_

---

**User**

Eres un especialista en programacion en python. Necesito que desarrolles una aplicacion considerando los siguientes criterios.

1.Tienes que abrir la ventana de un navegador controlado a traves de la libreria de playwright. La pagina es: https://api-seguridad.sunat.gob.pe/v1/clientessol/4f3b88b3-d9d6-402a-b85d-6a0bc857746a/oauth2/loginMenuSol?lang=es-PE&showDni=true&showLanguages=false&originalUrl=https://e-menu.sunat.gob.pe/cl-ti-itmenu/AutenticaMenuInternet.htm&state=rO0ABXNyABFqYXZhLnV0aWwuSGFzaE1hcAUH2sHDFmDRAwACRgAKbG9hZEZhY3RvckkACXRocmVzaG9sZHhwP0AAAAAAAAx3CAAAABAAAAADdAADZXhlcHQABnBhcmFtc3QASyomKiYvY2wtdGktaXRtZW51L01lbnVJbnRlcm5ldC5odG0mYjY0ZDI2YThiNWFmMDkxOTIzYjIzYjY0MDdhMWMxZGI0MWU3MzNhNnQABGV4ZWNweA== 

2. Ingresaras a la pagina y te vas a logear por lo que debes creare un fichero credenciales.json donde estara el ruc, usuario y contraseña.

El ruc esta en: <input type="text" placeholder="RUC" id="txtRuc" class="form-control" maxlength="11" value="" style="border-color: rgb(255, 64, 129);">

El usuario esta en: <input type="text" placeholder="RUC" id="txtRuc" class="form-control" maxlength="11" value="" style="border-color: rgb(255, 64, 129);">

La contraseña esta en: <input type="password" placeholder="Contraseña" id="txtContrasena" class="form-control" maxlength="12" onkeypress="return onPressEnter(this,event)" value="">

Crea un venv llamado venv_sunat e instala todas las librerias necesarias.


---

**Cursor**

Creando la aplicación Python con Playwright para automatizar el inicio de sesión en SUNAT. Revisando la estructura del proyecto y creando los archivos necesarios.


Creando la aplicación completa: los campos del formulario son `txtRuc`, `txtUsuario` y `txtContrasena`.


Comprobando si la página usa solo RUC y contraseña (sin campo usuario separado) y creando el venv.


La página muestra RUC y DNI; el flujo puede ser en dos pasos. Ajustando el script para esperar a que aparezcan los campos y manejando ambos flujos. Creando el venv e instalando dependencias.


Ajustando el script para soportar un flujo en dos pasos (primero RUC, luego usuario y contraseña), por si SUNAT lo usa:


Resumen de lo que quedó creado:

---

## Aplicación lista

### 1. **`credenciales.json`**
Archivo con tres claves: `ruc`, `usuario`, `contrasena`. Debes reemplazar los valores `TU_...` por tus datos reales de SUNAT SOL (RUC de 11 dígitos, usuario SOL y clave SOL).

### 2. **`login_sunat.py`**
Script que:
- Abre un navegador Chromium con **Playwright**
- Entra en la URL de login SUNAT Menú SOL que indicaste
- Lee las credenciales de `credenciales.json`
- Rellena **RUC** (`#txtRuc`), **Usuario** (`#txtUsuario`, si existe) y **Contraseña** (`#txtContrasena`)
- Soporta flujo en un solo paso (todos los campos a la vez) o en dos (primero RUC, luego “Continuar” y después usuario/contraseña)
- Pulsa “Iniciar sesión” (o envía el formulario) y deja la ventana abierta hasta que pulses Enter en la consola

### 3. **`requirements.txt`**
Incluye `playwright==1.49.0`.

### 4. **Entorno virtual `venv_sunat`**
- Creado en la carpeta del proyecto
- Con Playwright y Chromium instalados

### Cómo ejecutarlo

1. Edita `credenciales.json` y pon tu RUC, usuario SOL y contraseña.
2. En PowerShell, desde la carpeta del proyecto:

```powershell
.\venv_sunat\Scripts\Activate.ps1
python login_sunat.py
```

Si la página de SUNAT tiene otro flujo (por ejemplo, solo RUC y contraseña sin usuario, o pasos distintos), coméntalo y adaptamos los selectores o el orden de los pasos en `login_sunat.py`. En el README hay instrucciones y notas adicionales.

---

**User**

Perfecto, lo que debees hacer ahora es:

1. Ingresa donde dice Empresas, lo encuentras como: <h4>
                            <span class="hidden-xs">
                                <img src="/a/imagenes/sunatico01.png" class="imgIconoSunat1">&nbsp;&nbsp;
                            </span>
                            Empresas</h4>
2. Ingresar a Comprobante de pago: <span class="spanNivelDescripcion">Comprobantes de pago</span>
3. Cuando ingresas a Comprobante de pago se abre una lista y debes ingresar a SEE - SOL <span class="spanNivelDescripcion">SEE - SOL</span>
4. Cuando ingreses a SEE - SOL tambien se abre otra lista
5. Vas a Consultar Factura y Nota <span class="spanNivelDescripcion">Consultar Factura y Nota</span> Donde das click e ingresas
6. Vas a ir donde esta una casilla de Fecha de Inicio <input class="dijitReset dijitInputInner" data-dojo-attach-point="textbox,focusNode" autocomplete="off" name="fec_desde" type="text" tabindex="0" id="criterio.fec_desde" size="10" maxlength="10" value="" aria-invalid="false">    en este apartado pones lo que ingresa el usuario en la vista de tkinter que es un formato de dia/mes/año por ejemplo 01/01/2026
7. VVas a ir donde esta una casilla de Fecha de Fin <input class="dijitReset dijitInputInner" data-dojo-attach-point="textbox,focusNode" autocomplete="off" name="fec_hasta" type="text" tabindex="0" id="criterio.fec_hasta" size="10" maxlength="10" value="" aria-invalid="false">      en este apartado pones lo que ingresa el usuario en la vista de tkinter que es un formato de dia/mes/año por ejemplo 31/01/2026
8. Luego le das al boton Aceptar <span class="dijitReset dijitInline dijitIcon nextIcon" data-dojo-attach-point="iconNode"></span>


Importante que es la parte visual de la aplicacion. muestre primereo las credenciales y tenga para poner Fecha de Inicio y Fin los cuales el RPA usara para descargar las facturas. En esta parte visual debe tener un boton llamado proceder cuando recien le doy inicia el flujo. 

---

**Cursor**

Creando la interfaz Tkinter y extendiendo el flujo RPA para navegar hasta la consulta de facturas. Revisando el código actual:


La aplicación queda lista. Resumen de lo implementado:

---

## Cambios realizados

### 1. **Interfaz Tkinter**

- **Credenciales**: RUC, Usuario y Contraseña (se cargan desde `credenciales.json` si existe).
- **Fechas**: Fecha de Inicio y Fecha de Fin en formato `dd/mm/aaaa` (ej: 01/01/2026, 31/01/2026).
- **Área de estado**: Muestra el progreso del RPA.
- **Botón "Proceder"**: Al hacer clic se inicia todo el flujo (login + navegación + consulta).

### 2. **Flujo RPA**

1. Login en SUNAT con las credenciales.
2. Clic en **Empresas** (`h4` con texto "Empresas").
3. Clic en **Comprobantes de pago** (`span.spanNivelDescripcion`).
4. Clic en **SEE - SOL**.
5. Clic en **Consultar Factura y Nota**.
6. Relleno de **Fecha de Inicio** (`criterio.fec_desde`) con el valor de la interfaz.
7. Relleno de **Fecha de Fin** (`criterio.fec_hasta`).
8. Clic en **Aceptar** (botón con icono `nextIcon` o texto "Aceptar").

### 3. **Validaciones**

- Formato de fechas `dd/mm/aaaa`.
- Credenciales obligatorias antes de proceder.
- Las credenciales se guardan en `credenciales.json` al hacer clic en Proceder.

### Cómo ejecutar

```powershell
.\venv_sunat\Scripts\Activate.ps1
python login_sunat.py
```

Se abrirá la ventana de Tkinter. Comprueba las credenciales, introduce las fechas y pulsa **Proceder** para iniciar el flujo. El RPA se ejecuta en segundo plano para que la interfaz siga respondiendo.

---

**User**

Perfecto, ahora lo que debes hacer es:

1. Debes alejar la vista de la pagina a 60%.
2. Mover la linea de abajo hasta la derecha del todo(imagen adjunta) luego darle click en Descargar PDF de cada fila disponible imagen adjunta. "Descargar PDF" es el nombre del link.
3. Cuando le des click se descargar un pdf el cual lo guardaras en la carpeta FACTURAS(carpeta que vas a crear) este pdf se debe guardar considerando el año y mes que se puso inicialmente es decir por ejemplo si es usuario puso 01/01/2026 - 31/01/2026 y considerando la razon de la empresa entonces la carpeta seria FACTURA/2026/ENERO/ INVERSIONES GIANILU S.A.C.   Ahi dentro estaria los pdf.
4. Las facturas se deben nombrar considerando los siguientes criterios, por ejemplo @PDF-DOC-E00115420609802295.pdf con esta factura deberia ser E001_154_CINTA ESPUMA LINER AMARILLO 2 CM X 30 MTS. Del documento en mencion podemos ver la razon social de la empresa INVERSIONES GIANILU S.A.C. Lo encuentras en Señor(es):
5. Debes crear un excel que estara afuera es decir en FACTURAS que es una extraccion del contenido de la factura considerando estos criterios:.

FECHA	EMPRESA	CANTIDAD	DESCRIPCION	MONEDA	P_UNIT	TOTAL	URL

Continuando con el ejemplo esperaria el siguiente resultado

FECHA	EMPRESA	CANTIDAD	DESCRIPCION	MONEDA	P_UNIT	TOTAL	URL
20/01/2026	INVERSIONES GIANILU S.A.C.	30	CINTA ESPUMA LINER AMARILLO 2 CM X 30 MTS	SOLES	14	420	C:\Users\Ing.Robinzon\Desktop\Hackaton\FACTURAS\2026\ENERO\INVERSIONES GIANILU S.A.C

Donde el url sea un hipervinculo que con un click pueda abrir el pdf.

---

**Cursor**

Voy a implementar la descarga automática de todos los “Descargar PDF”, guardarlos con la estructura de carpetas solicitada, renombrarlos según la serie/número y descripción del PDF, y generar el Excel con hipervínculos. Primero revisaré el PDF de ejemplo para ajustar la extracción.

Ahora agregaré las dependencias para leer PDFs y generar el Excel, y después probaré que el script arranca y llega a la parte de descarga (sin obligarte a loguear ahora).

## Listo: descarga masiva + renombrado + Excel

### ✅ Lo que ya hace `login_sunat.py`
- **Zoom al 60%** en la pantalla de resultados.
- **Mueve el scroll horizontal al extremo derecho** para que se vean los links **“Descargar PDF”**.
- **Hace clic en “Descargar PDF” en cada fila** y guarda todos los PDFs.
- **Crea la estructura de carpetas**:

`FACTURAS/<AÑO>/<MES>/<EMPRESA>/`

Ejemplo: `FACTURAS/2026/ENERO/INVERSIONES GIANILU S.A.C./`

- **Renombra cada PDF** con el formato:

`E001_154_CINTA_ESPUMA_LINER_AMARILLO_2_CM_X_30_MTS.pdf`

(usa `E001-154` y la **primera descripción** extraída del PDF).

- **Genera un Excel** en `FACTURAS/EXTRACCION_FACTURAS.xlsx` con columnas:

`FECHA | EMPRESA | CANTIDAD | DESCRIPCION | MONEDA | P_UNIT | TOTAL | URL`

y la columna **URL es un hipervínculo** al PDF local.

### 🧩 Dependencias agregadas
Actualicé `requirements.txt` y ya están instaladas en `venv_sunat`:
- `openpyxl` (Excel)
- `pypdf` (lectura/extracción del PDF)

### ▶️ Cómo ejecutarlo
En PowerShell:

```powershell
.\venv_sunat\Scripts\Activate.ps1
python login_sunat.py
```

En la ventana Tkinter:
- Revisa credenciales
- Ingresa **Fecha Inicio/Fin**
- Clic en **Proceder**

Al terminar tendrás:
- PDFs guardados y renombrados en `FACTURAS\...`
- Excel `FACTURAS\EXTRACCION_FACTURAS.xlsx` con hipervínculos

Si en tu SUNAT el link “Descargar PDF” está dentro de un **iframe distinto** al de fechas/resultados, dímelo (o pásame el HTML de esa zona) y ajusto el “target” para encontrar los links siempre.

---

**User**

Perfecto, ahora necesito que realices los sigueintes cambios:

1. Al nivel de interfaz de usuario las fechas debe ser como un calendario que cuando de click se despliegue un tipo calendario.
2. Necesito que el excel contenga informacion de todas las carpetas que se crean y este segmentado por año es decir, si la carpeta FACTURAS tiene 2025,2026 el excel deberia tener 2 hojas 1 por cada año.
3. Necesito que en la interfaz de usuario el usuario tenga la opcion de poder colocar meses de año por ejemplo si coloca enero 2025 y diciembre 2025. El programa debe empezar mes por mes completo, ir descargando campo por campo y luego generar el excel. En este ejemplo iniciara con 01/01/2025 - 31/01/2025, luego que descarga tiene que ir a Consultar Factura y Nota para ingresar la fecha del siguiente mes y de esata forma hasta culminar la fecha que puso el usuario<span class="spanNivelDescripcion">Consultar Factura y Nota</span>
4. Necesito que el excel tenga una hoja llamada analisis que sea una tabla dinamica que este por año, mes, cliente, producto y total. Que se genere un gráfico que responda la siguiente pregunta ¿Que cliente compra mas por mes y año? ¿Cual es la descripcion que mas se vende por mes y año?

---

**Cursor**

Voy a hacer 3 cambios grandes: reemplazar las fechas por un calendario (widget tipo date-picker), permitir un rango por **mes/año** para que el bot procese mes a mes, y rearmar el Excel con **una hoja por año** más una hoja **ANALISIS** con agregaciones y gráficos. Empezaré revisando las partes actuales de UI y Excel para modificarlas con el menor riesgo.

## Cambios aplicados

### 1) Calendario en la interfaz (date-picker)
- Ya no se ingresan fechas “a mano”.
- En la UI ahora tienes:
  - **Mes/Año Inicio** (combo + año)
  - **Mes/Año Fin** (combo + año)
  - Un **calendario desplegable** (tipo calendario) como referencia visual (widget `DateEntry` de `tkcalendar`).

Además agregué la dependencia `tkcalendar` en `requirements.txt` y quedó instalada en `venv_sunat`.

---

### 2) Excel con **hojas por año** (toma todas las carpetas generadas por el proceso)
- El Excel ahora se genera como `FACTURAS\EXTRACCION_FACTURAS.xlsx`.
- Contiene:
  - **Una hoja por año**: `2025`, `2026`, etc. (con las columnas: FECHA, EMPRESA, CANTIDAD, DESCRIPCION, MONEDA, P_UNIT, TOTAL, URL).
  - La columna **URL** queda como **hipervínculo** al PDF.

---

### 3) Procesamiento **mes por mes** (enero→diciembre, etc.)
- El flujo ahora soporta: “**ENERO 2025** a **DICIEMBRE 2025**”.
- Automáticamente hace:
  - 01/01/2025–31/01/2025 → descarga
  - vuelve a **“Consultar Factura y Nota”**
  - 01/02/2025–28/02/2025 → descarga
  - … hasta el último mes del rango.

Esto está implementado en la nueva función `ejecutar_rpa_por_meses(...)` y la UI ya la llama cuando presionas **Proceder**.

---

### 4) Hoja **ANALISIS** (tipo “tabla dinámica” + gráficos)
- Se crea una hoja `ANALISIS` con:
  - **Top cliente por mes y año** (y su total)
  - **Top descripción por mes y año** (y su total)
  - **2 gráficos de barras** (uno para cliente top, otro para producto top)

> Nota: Excel “PivotTable” real (objeto de tabla dinámica nativo) es muy limitado de generar con `openpyxl` y suele requerir estructuras internas complejas; por eso lo resolví como **resumen tipo tabla dinámica + gráficos**, que responde exactamente las 2 preguntas.

---

## Cómo ejecutar

```powershell
.\venv_sunat\Scripts\Activate.ps1
python login_sunat.py
```

En la ventana:
- Define **Mes/Año inicio** y **Mes/Año fin**
- Clic en **Proceder**

Cuando termine tendrás:
- PDFs en `FACTURAS\<AÑO>\<MES>\<EMPRESA>\...`
- Excel `FACTURAS\EXTRACCION_FACTURAS.xlsx` con hojas por año + `ANALISIS`

Si quieres que el Excel “incluya también PDFs ya existentes en FACTURAS aunque no se hayan descargado en esta corrida”, lo ajusto para que además escanee todo `FACTURAS\**\*.pdf` y reconstruya el Excel completo.

---

**User**

Cuando se termine de procesar, que salga una ventana emergente indicando se termino el proceso. ¿Pregunta desea cerrar el navegador? Si le doy si se cierra si le doy no se mantiene abierto. Independiente del resultado si o no, se debe abrir la carpeta (facturas) y el excel.

---

**Cursor**

### Cambios implementados

- **Al finalizar todo el procesamiento (todos los meses, todas las facturas):**
  - Se genera el Excel `EXTRACCION_FACTURAS.xlsx` como antes.
  - **Se abre automáticamente**:
    - La carpeta `FACTURAS`.
    - El archivo de Excel `EXTRACCION_FACTURAS.xlsx`.

- **Luego aparece una ventana emergente (messagebox):**

  Texto:
  > “Se terminó el proceso.  
  > ¿Desea cerrar el navegador?”

  - Si eliges **Sí**:
    - Se cierra el navegador de Playwright.
  - Si eliges **No**:
    - El navegador se mantiene abierto y el programa queda esperando a que tú lo cierres manualmente (usa `mantener_navegador_abierto`).

No necesitas cambiar nada en tu forma de uso: ejecutas `login_sunat.py`, configuras rango de meses/años y pulsas **Proceder**; al terminar verás el popup y se abrirán automáticamente la carpeta y el Excel.

