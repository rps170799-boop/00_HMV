# Bienvenido a HMV Tools 🚀

**HMV Tools** es una suite de herramientas (Add-ins) desarrollada en C# para la API de Autodesk Revit. Su objetivo es automatizar tareas repetitivas, auditar modelos y estandarizar la información de los proyectos.

---

## 📦 Instalación
Para instalar las herramientas en tu equipo:
1. Descarga la última versión del instalador `.msi` desde la sección de **Releases** en GitHub.
2. Cierra Revit.
3. Ejecuta el instalador y sigue las instrucciones.
4. Al abrir Revit, encontrarás la nueva pestaña **HMV Tools** en la cinta de opciones (Ribbon).

---

## 🛠️ Resumen de Módulos

Nuestra cinta de opciones está dividida en 4 paneles principales para facilitar tu flujo de trabajo:

### 1. DWG
Herramientas para la gestión y conversión de archivos CAD importados.
* **DWG Convert:** Convierte líneas y textos de DWG a los estándares de HMV.
* **3D DWG to Shape:** Extrae sólidos y mallas de un DWG 3D a un DirectShape nativo.

### 2. Family Control Tools
Scripts para edición masiva y posicionamiento espacial de familias.
* **Ductos Editor:** Configurador visual interactivo de tuberías USD.
* **Refresh Z to Floor:** Proyecta elementos hacia la cara superior de un suelo vinculado.
* **Multi InstParam Editor:** Edita parámetros de múltiples instancias en una sola transacción.
* **Foundation Control:** Ajusta la elevación de cimentaciones según niveles topográficos.

### 3. Annotation Tools
Automatización de tareas de documentación y etiquetado.
* **Grid/Level Extent:** Alterna masivamente entre extensión 2D y 3D.
* **Topo to Lines:** Extrae curvas de nivel de un vínculo a líneas de detalle.
* **Align Spot Elevations:** Alinea cotas de elevación a un eje común.

### 4. Audit
Herramientas para mantener la salud y el estándar del modelo.
* **Text / Dim Audit:** Estandariza fuentes, tamaños y nombres de tipos.
* **View / Sheet Audit:** Renombra vistas y planos detectando duplicados.
* **Family Audit:** Compara versiones de familias contra la nube (ADC) usando hashes SHA-256.

---

> Usa el menú superior o la barra de búsqueda para ver el detalle y funcionamiento de cada comando específico.