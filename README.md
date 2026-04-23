# Certificate Generator · Sheets to Docs

Google Apps Script que automatiza la generación de certificados en Google Docs
a partir de datos en una hoja de cálculo, usando múltiples plantillas almacenadas
en Google Drive.

## ✨ Características

- Genera certificados personalizados en Google Docs desde una hoja de cálculo.
- Soporta múltiples plantillas con selección dinámica por fila.
- Reemplaza marcadores dinámicos (`{{campo}}`) con los datos de cada fila.
- Nombra los archivos generados de forma estructurada: `Nombre - Tipo de plantilla`.
- Convierte automáticamente el nombre del estudiante a mayúsculas en el documento.
- Menú integrado en Google Sheets para ejecución con un clic.
- Caché de plantillas para optimizar rendimiento en lotes grandes.
- Manejo de errores por fila: si una falla, las demás siguen generándose.
- Configuración segura mediante Propiedades del Script (IDs no expuestos en el código).

## 🗂️ Estructura del proyecto

```
certificate-generator-sheets/
├── Codigo.gs         # Script principal
├── README.md         # Este archivo
├── LICENSE           # Licencia MIT
└── .gitignore        # Archivos ignorados por Git
```

## 🚀 Instalación

### 1. Preparar Google Drive

Crea dos carpetas en Google Drive:

- **Carpeta de plantillas**: contendrá los archivos de Google Docs que servirán como plantilla.
- **Carpeta de destino**: aquí se guardarán los certificados generados.

Copia el **ID de cada carpeta** desde la URL de Drive:

```
https://drive.google.com/drive/folders/ESTE_ES_EL_ID
```

### 2. Preparar las plantillas

Cada plantilla debe ser un Google Docs con marcadores del tipo `{{nombre_campo}}`
donde quieras insertar datos dinámicos.

Ejemplo de contenido de plantilla:

```
Se certifica que {{nombre_estudiante}}
ha completado el programa {{programa}}
con fecha {{fecha}}.
```

### 3. Preparar la hoja de cálculo

La primera fila debe contener los encabezados. Son obligatorios:

- `nombre_estudiante`: nombre completo del estudiante.
- `Plantilla`: nombre del archivo de plantilla a usar (debe existir en la carpeta de plantillas).

Cualquier otra columna cuyo encabezado coincida con un marcador `{{campo}}` de la
plantilla será reemplazada automáticamente.

### 4. Instalar el script

1. Abre la hoja de cálculo.
2. Ve a **Extensiones → Apps Script**.
3. Borra el contenido por defecto y pega el código de `Codigo.gs`.
4. Guarda el proyecto (`Ctrl + S` o `Cmd + S`).

### 5. Configurar los IDs de carpeta

1. En el editor de Apps Script, haz clic en el icono de **engranaje ⚙️**
   (**Configuración del proyecto**) en la barra lateral izquierda.
2. Baja hasta la sección **Propiedades del script**.
3. Haz clic en **Añadir propiedad del script** y crea estas dos propiedades:

| Propiedad | Valor |
|---|---|
| `CARPETA_PLANTILLAS_ID` | ID de la carpeta de plantillas |
| `CARPETA_DESTINO_ID` | ID de la carpeta de destino |

4. Haz clic en **Guardar propiedades del script**.

### 6. Autorizar el script

1. Vuelve a la hoja de cálculo y recárgala.
2. En el menú superior aparecerá **📄 Documentos**.
3. La primera vez que uses una opción, Google pedirá autorización.
4. Acepta los permisos solicitados.

## 💻 Uso

En la hoja de cálculo, usa el menú **📄 Documentos**:

- **Generar certificados**: procesa todas las filas con datos.
- **Certificado (fila actual)**: procesa solo la fila seleccionada.

Los archivos generados se guardarán en la carpeta de destino con el formato:

```
NOMBRE ESTUDIANTE - TIPO DE PLANTILLA
```

## 🧩 Ejemplo de uso

**Hoja de cálculo:**

| nombre_estudiante | programa | fecha | Plantilla |
|---|---|---|---|
| Juan Pérez | Sistemas | 2026-04-23 | 3. PLANTILLA TECNICA PROFESIONAL EN SISTEMAS |

**Archivo generado:**

```
Juan Pérez - TECNICA PROFESIONAL EN SISTEMAS
```

**Contenido del archivo (después del reemplazo):**

```
Se certifica que JUAN PÉREZ
ha completado el programa Sistemas
con fecha 2026-04-23.
```

## 🛠️ Tecnologías

- **Google Apps Script** — Runtime de JavaScript para el ecosistema Google.
- **Google Sheets API** — Lectura de datos de la hoja de cálculo.
- **Google Drive API** — Gestión de archivos y carpetas.
- **Google Docs API** — Manipulación del contenido de los certificados.

## 📋 Requisitos

- Cuenta de Google con acceso a Sheets, Docs y Drive.
- Permisos de edición sobre las carpetas de plantillas y destino.

## 🧪 Buenas prácticas implementadas

- **Separación de configuración y código**: los IDs sensibles no se almacenan en el código fuente, sino en Propiedades del Script, permitiendo publicar el repositorio sin exponer información interna.
- **Normalización de nombres**: la búsqueda de plantillas es tolerante a variaciones de mayúsculas, tildes y espacios.
- **Manejo robusto de errores**: cada fila se procesa de forma aislada, de modo que un fallo puntual no interrumpe el lote completo.
- **Optimización con caché**: las plantillas encontradas se guardan en memoria durante la ejecución para evitar búsquedas repetidas en Drive.

## 📄 Licencia

Este proyecto está bajo la licencia MIT. Consulta el archivo [LICENSE](./LICENSE) para más detalles.

## 👤 Autor

**Brahan Julián García Solórzano**

- GitHub: [@bjgarcia20](https://github.com/bjgarcia20)

---

⭐ Si este proyecto te resultó útil, considera darle una estrella en GitHub.
