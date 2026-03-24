# 📋 Sistema de Registro de Asistencia Sindical
### Sindicato de Trabajadores — SLIM N°3

Sistema automatizado de control de asistencia a asambleas sindicales, desarrollado en **Google Apps Script** sobre Google Sheets. Valida RUTs chilenos, detecta duplicados, protege registros y envía confirmaciones por correo electrónico con diseño HTML.

---

## 🚀 Características principales

- ✅ **Validación de RUT chileno** con cálculo de dígito verificador
- 🔍 **Detección de duplicados** en tiempo real al ingresar un RUT
- 🔒 **Protección de registros** para evitar edición accidental de RUTs ya confirmados
- 📧 **Envío automático de correo HTML** con código de verificación único
- 🧹 **Limpieza masiva de datos** con confirmación previa
- ⚙️ **Arquitectura orquestada** con pasos secuenciales y logging detallado

---

## 📁 Estructura del proyecto

```
├── Orquestador.gs          # Función maestra que coordina todo el flujo
├── validacionRuts.gs       # Lógica de validación del RUT chileno
├── SistemaAsistencia.gs    # Generación de HTML y envío de correos
└── limpiarDatos.gs         # Limpieza de datos y gestión de protecciones
```

---

## 🔄 Flujo de ejecución

El sistema se activa automáticamente al editar la **columna A (RUT)** de cualquier fila (excepto el encabezado):

```
Edición en columna A
        │
        ▼
┌───────────────────────┐
│  PASO 1               │  ¿Está la protección activa?
│  Verificar protección │  Si el RUT ya existe → BLOQUEAR y restaurar valor
└───────────┬───────────┘
            │ ✅ Permitido
            ▼
┌───────────────────────┐
│  PASO 2               │  ¿El RUT ya existe en otra fila?
│  Verificar duplicados │  Si hay duplicado → LIMPIAR celda y alertar
└───────────┬───────────┘
            │ ✅ Único
            ▼
┌───────────────────────┐
│  PASO 3               │  ¿Es un RUT chileno válido?
│  Validar formato RUT  │  Escribe "VALIDO" o "NO VALIDO" en columna B
└───────────┬───────────┘
            │ ✅ Válido
            ▼
┌───────────────────────┐
│  PASO 4               │  Genera código único, protege celda RUT,
│  Enviar correo HTML   │  envía confirmación y actualiza columna H
└───────────────────────┘
```

---

## 🗂️ Estructura de la hoja de cálculo

| Columna | Letra | Contenido |
|---------|-------|-----------|
| 1 | A | RUT |
| 2 | B | Validación de RUT (`VALIDO` / `NO VALIDO`) |
| 3 | C | Nombre |
| 8 | H | Validación de correo / Estado de envío |
| 10 | J | Correo electrónico |
| 11 | K | Código de verificación único |
| 12 | L | Asamblea (mes) |

---

## ⚙️ Instalación y configuración

### 1. Copiar los scripts

Abre tu Google Sheet y ve a **Extensiones → Apps Script**. Crea un archivo `.gs` por cada archivo de este repositorio y pega el contenido correspondiente.

### 2. Configurar el trigger

Solo debe existir **un único trigger** activo de tipo `onEdit`:

1. En Apps Script, ve a **Activadores (Triggers)**
2. Crea un nuevo activador:
   - Función: `orquestadorMaestro`
   - Tipo de evento: `Al editar (onEdit)`
3. Asegúrate de **eliminar** cualquier trigger previo de `onEdit` u `enviarCorreoHTML`

### 3. Permisos requeridos

Al ejecutar por primera vez, Google solicitará autorizar los siguientes permisos:
- 📧 Envío de correos (`MailApp`)
- 📊 Lectura y escritura en hojas de cálculo
- 🔐 Almacenamiento de propiedades del script (`PropertiesService`)

---

## 🛠️ Funciones disponibles para ejecución manual

| Función | Archivo | Descripción |
|--------|---------|-------------|
| `validateAllRutsInColumnA()` | `validacionRuts.gs` | Valida todos los RUTs de la columna A de una vez |
| `limpiarRutCodigoYObservacion()` | `limpiarDatos.gs` | Limpia datos seleccionados con confirmación previa |
| `permitirEdicionRUT()` | `limpiarDatos.gs` | Desactiva temporalmente la protección de RUTs |
| `bloquearEdicionRUT()` | `limpiarDatos.gs` | Reactiva la protección de RUTs |
| `verificarEstadoProteccion()` | `limpiarDatos.gs` | Muestra si la protección está activa o no |
| `contarProteccionesRUT()` | `limpiarDatos.gs` | Cuenta celdas protegidas en columna RUT |
| `probarOrquestador()` | `Orquestador.gs` | Simula manualmente el flujo para una fila de prueba |

---

## 🔒 Sistema de protección de RUTs

Una vez que un RUT es procesado exitosamente (correo enviado), su celda queda **protegida contra edición**. Esto evita modificaciones accidentales del registro.

Para editar un RUT protegido:

```
1. Ir a Apps Script
2. Ejecutar: permitirEdicionRUT()
3. Realizar la edición en la hoja
4. Ejecutar: bloquearEdicionRUT()
```

> ⚠️ No olvides reactivar la protección después de realizar los cambios.

---

## 📧 Correo de confirmación

El sistema genera y envía un correo HTML con:

- ✅ Encabezado con gradiente azul institucional
- 👤 Nombre personalizado del afiliado
- 🪪 RUT del socio
- 🔑 Código de verificación único de 11 caracteres (alfanumérico)
- 📅 Mes de la asamblea
- ⚠️ Instrucciones para apelación de multa
- 🔗 Enlace directo al portal de afiliados: [sindicatoslim3.com](https://www.sindicatoslim3.com/aplicaciones/app-login)

---

## 🧮 Algoritmo de validación de RUT chileno

La validación sigue el estándar oficial chileno:

1. Se separa el **cuerpo numérico** del **dígito verificador** (último carácter)
2. Se aplica el **algoritmo módulo 11** con multiplicadores cíclicos del 2 al 7
3. Se compara el dígito calculado con el ingresado
4. Acepta formatos: `178592130`, `17.859.213-0`, `17859213-0`, con o sin puntos y guión

---

## 🐛 Registro de errores

El sistema registra errores en la **columna H** de la fila afectada con el prefijo `ERROR:`, y adicionalmente en el **Logger de Apps Script** (`Ver → Registros`).

---

## 🤝 Contribuciones

Este proyecto es de uso interno del **Dpto. de Comunicaciones — SLIM N°3**. Para proponer mejoras, abrir un _Issue_ o un _Pull Request_ en este repositorio.

---

## 📄 Licencia

Uso interno — Sindicato Libre de Trabajadores de Metro S.A. N°3. Todos los derechos reservados.
