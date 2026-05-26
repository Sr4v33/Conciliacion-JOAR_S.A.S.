# Rectificadora JOAR — Automatización de Conciliación DIAN vs Siigo

Automatización desarrollada en Python para el proceso de conciliación tributaria entre reportes descargados desde la DIAN y libros oficiales exportados desde Siigo.
El sistema permite identificar automáticamente discrepancias entre ambos sistemas, generar reportes consolidados y notificar los resultados mediante correo electrónico.

---

# Objetivo

Automatizar el proceso de validación y conciliación entre:

- Reportes de facturas recibidas descargados desde la DIAN.
- Libros oficiales exportados desde Siigo.

La solución permite:

- Comparación automática de registros.
- Detección de facturas faltantes.
- Generación de reportes Excel.
- Consolidación de métricas de conciliación.
- Envío automático de correos.
- Ejecución local mediante ejecutable.

---

# Tecnologías utilizadas

- Python 3.x
- pandas
- openpyxl
- PyInstaller

---

# Estructura del proyecto

```bash
/
│
├── conciliacion-dian-siigo/
│   ├── build.spec
│   ├── conciliador.py
│   ├── config.ini
│   ├── enviador_correo.py
│   ├── generador_excel.py
│   ├── main.py
│   └── utils.py
│
├── conciliacion-siigo-dian/
│   ├── build.spec
│   ├── conciliador.py
│   ├── config.ini
│   ├── enviador_correo.py
│   ├── generador_excel.py
│   ├── main.py
│   └── utils.py
│
├── ejecutables/
│   │
│   ├── Dian_Siigo/
│   │   ├── ConciliacionDIAN_Siigo_JOAR.exe
│   │   └── config.ini
│   │
│   └── Siigo_Dian/
│       ├── ConciliacionSiigo_DIAN_JOAR.exe
│       └── config.ini
│
├── .gitattributes
└── requirements.txt
```

---

# Funcionamiento general

## 1. Entrada de archivos

El sistema trabaja sobre archivos previamente descargados desde:

- Portal DIAN.
- ERP Siigo.

Los archivos deben conservar la estructura original exportada por cada plataforma.

---

## 2. Procesamiento DIAN

La automatización aplica filtros automáticos sobre el archivo DIAN:

### Tipos de documento válidos

- Factura electrónica
- Documento equivalente POS
- Nota de crédito electrónica

### Grupo válido

- `Recibido`

El resultado corresponde al universo de facturas válidas para conciliación.

---

## 3. Procesamiento Siigo

El sistema procesa el libro oficial de compras exportado desde Siigo.

A partir del campo:

```text
Factura proveedor
```

se extrae automáticamente el folio numérico.

### Ejemplo

```text
FE-2363 → 2363
```

---

## 4. Conciliación automática

El sistema compara:

- Folio DIAN
vs
- Factura proveedor extraída desde Siigo

### Resultados posibles

#### Coincidencia encontrada

- Registro resaltado visualmente en amarillo.

#### Factura presente en DIAN pero no en Siigo

- Registro marcado en rojo.
- Registro enviado a hoja `Faltantes`.

---

## 5. Generación de reporte

El sistema genera automáticamente un archivo Excel con:

### Hoja Resumen

Incluye:

- Fecha y hora de ejecución.
- Período conciliado.
- Total facturas DIAN.
- Total facturas Siigo.
- Facturas conciliadas.
- Facturas faltantes.
- Valor acumulado faltantes.

### Hoja Faltantes

Incluye:

- Registros sin coincidencia.
- Consolidado total de valores faltantes.

---

## 6. Envío automático de correo

Una vez finalizado el proceso:

- Se adjunta el reporte generado.
- Se envía automáticamente al correo configurado.

Porque generar Exceles y enviarlos por correo sigue siendo la columna vertebral de media economía latinoamericana. Civilización digital de altísima sofisticación.

---

# Ejecución del proyecto

## 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

---

## 2. Configurar parámetros

Editar el archivo:

```text
config.ini
```

Incluyendo:

- rutas de archivos,
- credenciales de correo,
- destinatarios,
- configuración general.

---

## 3. Ejecutar sistema

### Desde Python

```bash
python main.py
```

### Desde ejecutable

Ejecutar:

```text
ConciliacionDIAN_Siigo_JOAR.exe
```

o

```text
ConciliacionSiigo_DIAN_JOAR.exe
```

según el flujo requerido.

---

# Resultado esperado

El sistema genera:

- Reporte consolidado Excel.
- Hoja de faltantes.
- Métricas de conciliación.
- Correo automático con resultados.

---

# Consideraciones

- La solución opera localmente.
- No existe integración directa mediante API con DIAN o Siigo.
- La automatización depende de archivos previamente descargados.
- La estructura de los archivos debe mantenerse consistente.

---

# Posibles mejoras futuras

- Integración directa con plataformas externas.
- Programación automática periódica.
- Dashboard web.
- Base de datos de auditoría.
- Panel administrativo.

---

# Contexto académico

Proyecto desarrollado dentro del curso de Sistemas de Información, con enfoque en transformación digital y automatización aplicada a contextos reales (Rectificadora JOAR S.A.S.).

---

# Autores

- Jose Andrés Mendoza Hernández
- Juan Sebastián Rave Martínez
