### Determinación del propósito
El sistema **"Horizonte"** (específicamente los módulos `HTraspasos` y `HLiquidaciones`) es una solución empresarial de **Gestión Documental y Procesamiento de Traspasos de Pensiones (AFP)**. Su función principal es la automatización de la captura, digitalización, validación y exportación de solicitudes de afiliación y traspaso de fondos previsionales.

### Problemas empresariales que resuelve
*   **Digitalización Masiva:** Gestión de grandes volúmenes de documentos físicos (solicitudes, DNI, boletas de pago) mediante escaneo ISIS/VRS.
*   **Validación de Datos Previsionales:** Automatización de la verificación del **CUSPP** (Código Único del Sistema Privado de Pensiones) y datos del empleador.
*   **Trazabilidad por Lotes:** Agrupación de documentos en "Lotes" para un control de calidad y flujo de trabajo industrializado.
*   **Exportación Normativa:** Generación de archivos planos (`.PRO`) para el intercambio de datos con las AFPs (Horizonte, Profuturo, Integra) y entes reguladores.

### Áreas de negocio involucradas
*   **Operaciones/Back-office:** Procesamiento de la carga documental.
*   **RRHH / Previsión Social:** Gestión de afiliaciones y antecedentes de aportes.
*   **Logística Documental:** Almacenamiento físico y digital de imágenes.

---

## 2. TECNOLOGÍAS DETECTADAS

El sistema presenta un ecosistema **híbrido** típico de finales de los 90 y principios de los 2000. Aunque el repositorio se etiqueta como "VFP", el código fuente analizado corresponde principalmente a **Visual Basic 6.0**, interactuando con bases de datos SQL Server y archivos VFP (`.dbf`).

| Categoría | Tecnologías Identificadas |
| :--- | :--- |
| **Lenguaje Principal** | Visual Basic 6.0 (Estructura `.frm`, `.cls`, `.bas`). |
| **Base de Datos** | SQL Server (vía OLEDB) y tablas locales Visual FoxPro (`.dbf`). |
| **Drivers de Conexión** | ADO 2.x (ActiveX Data Objects), `SQLOLEDB`. |
| **Escaneo y Gráficos** | Kodak/Wang Imaging (`IMGEDIT.OCX`), Kofax Scan (`kscn3032.ocx`), LeadTools Barcode. |
| **Herramientas de Reporte** | Crystal Reports (`.rpt`) y DataReport Designers (`.Dsr`). |
| **Componentes COM/OLE** | `MSDATGRD.OCX` (Grids), `COMCT232.OCX` (Common Controls). |
| **Automatización** | Integración con Excel para reportes de inconsistencias (`.xls`). |
| **Infraestructura** | Scripts de Windows (`.bat`) para gestión de repositorios. |

---

## 3. ARQUITECTURA GENERAL

### Descripción de Capas
El sistema sigue un patrón de **dos capas (Client-Server)** con una separación incipiente de lógica:
1.  **Capa de Presentación (UI):** Formularios densos que contienen gran parte de la lógica de validación de negocio (ej. validación de CUSPP dentro de `frmAfiliacion.frm`).
2.  **Capa de Datos/Negocio:** 
    *   Uso de clases de conexión (`cConexion.cls`, `clsDataSQL.cls`) para centralizar el acceso a SQL Server.
    *   Lógica persistida en **Stored Procedures** (`pGetSecuAfi`, `pGrabarTRuc`).
    *   Uso de tablas temporales locales (`tempo.dbf`) para procesos rápidos de validación.

### Estructura y Acoplamiento
*   **Monolítico Cliente-Servidor:** El ejecutable debe estar instalado en cada estación de trabajo.
*   **Alto Acoplamiento:** La interfaz de usuario está íntimamente ligada a la estructura de la base de datos (uso extensivo de `Recordsets` vinculados directamente a controles).

---

## 4. MÓDULOS FUNCIONALES

| Módulo | Función Principal |
| :--- | :--- |
| **Digitalización (Scan)** | Interfaz con scanners industriales, control de DPI, rotación y limpieza de imágenes TIF. |
| **Asignación de Lotes** | Clasificación de carpetas de imágenes en unidades de trabajo (Batch/Mes/Año). |
| **Captura/Afiliación** | Digitación de datos del afiliado, empleador y antecedentes previsionales con visualización simultánea de la imagen. |
| **Validación D2** | Módulo de "Segunda Digitación" para garantizar que los datos ingresados por dos operadores coincidan (Control de Calidad). |
| **Reportes y Cruces** | Generación de reportes de errores por asesor, cruce de datos con RENIEC o empleadores (`TRUC`). |
| **Exportación** | Generación de archivos `.PRO` con formato de ancho fijo para las AFPs. |

---

## 5. MODELO DE DATOS INFERIDO

### Tablas Principales
*   **Maestras:**
    *   `Ubigeo`: Departamentos, provincias y distritos.
    *   `TAFP`: Catálogo de administradoras de fondos.
    *   `TPromotor`: Datos de los agentes de venta.
    *   `TRUC`: Maestro de empleadores validado.
*   **Transaccionales:**
    *   `Solicitud`: Cabecera de la afiliación o traspaso.
    *   `Empleadores`: Datos de las empresas vinculadas a la solicitud.
    *   `TB_Lotes`: Control de carga y estado de los lotes.
    *   `_Imagenes`: Metadatos de los archivos digitalizados (ID, Ruta, Posición).
*   **Históricas:**
    *   `TB_HISTORICO`: Registros de lotes ya procesados/importados.

---

## 6. SEGURIDAD Y AUDITORÍA

*   **Autenticación:** Basada en usuario y contraseña (`frmLogin.frm`) con validación contra tabla de usuarios.
*   **Validaciones de Integridad:** El código implementa algoritmos de **Módulo 11** para validar RUCs y validaciones complejas de longitud y formato para CUSPP.
*   **Auditoría de Procesos:** Campos en las tablas principales como `codusu`, `fecdig`, `horini`, `horfin` para rastrear quién y cuándo procesó cada registro.
*   **Perfiles:** Se detecta una jerarquía básica basada en prefijos de usuario (ej. "SUP" para Supervisores con acceso a corrección de datos generales).

---

## 7. RIESGOS Y DEUDA TÉCNICA

*   **Dependencias Obsoletas:** El sistema depende críticamente de controles OCX de 32 bits (Kodak Imaging) que no son compatibles de forma nativa con sistemas operativos modernos de 64 bits.
*   **Hardcoding de Rutas:** Configuraciones en archivos `.ini` con rutas físicas rígidas (ej. `C:\Horizonte\imagenes`) que dificultan la escalabilidad.
*   **Riesgo de Migración:** La lógica de negocio está fragmentada entre Stored Procedures en SQL y código en los eventos de los formularios, lo que complica una migración "as-is".
*   **Conexiones Inseguras:** Uso de usuario `sa` y contraseñas en texto claro o configuraciones locales (`Arccon.cnx`).

---

## 8. ESTRATEGIA DE MODERNIZACIÓN SUGERIDA

**Complejidad:** Alta | **Prioridad:** Crítica (debido a obsolescencia de componentes de imagen).

### Fase 1: Estabilización y Capa de Datos
*   Migrar toda la lógica de los archivos `.frm` a **Web APIs (ASP.NET Core o Node.js)**.
*   Centralizar la base de datos eliminando dependencias de archivos `.dbf` locales.

### Fase 2: Modernización del Frontend
*   Desarrollar una aplicación web (React o Angular).
*   **Sustitución de Imaging:** Reemplazar los OCX de escaneo por librerías web compatibles con estándares TWAIN/WIA o servicios de almacenamiento en la nube (AWS S3 / Azure Blob Storage).

### Fase 3: Integración y OCR
*   Implementar **OCR (Reconocimiento Óptico de Caracteres)** para automatizar la digitación que hoy es manual, reduciendo el error humano en campos como CUSPP y DNI.

**Estrategia recomendada:** *Strangler Fig Pattern*. Reemplazar módulos uno a uno (empezando por Reportes y Consultas, luego Captura) hasta que el monolito VB6 desaparezca por completo.
< 2018-01-05T16:44:00 --> 
< 2014-01-01T22:48:00 --> 
