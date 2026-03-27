# Bienvenido 

**HMV Tools** es una suite de herramientas (Add-ins) personalizadas desarrolladas en C# para la API de Autodesk Revit. Su objetivo es automatizar tareas repetitivas, auditar modelos y estandarizar la información de los proyectos.

---

## Instalación
Para instalar las herramientas en tu equipo:
<br>
1. Descarga la última versión del instalador `.msi` desde el entorno, ruta del archivo debe ser asignada por el coordinador BIM.
<br>
[Ir a la carpeta en ACC](https://acc.autodesk.com/docs/files/projects/992dbb1c-6004-472d-a416-934fff1806d1?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.DX2kA7XYTIG91f0G8eY18g&viewModel=detail&moduleId=folders){ .md-button }
<br>
> **Nota:** El acceso a esta carpeta puede estar restringido.
<br>
2. Cierra Revit.
<br>
3. Ejecuta el instalador y sigue las instrucciones.
<br>
4. Al abrir Revit, encontrarás la nueva pestaña **HMV Tools** en la cinta de opciones (Ribbon).

---

## Resumen de Módulos

La cinta de opciones está dividida en 4 paneles principales para facilitar el flujo de trabajo.
<br>
En producción significa que la herramienta está en fase de desarrollo y aún no cumple su proposito en totalidad.

### 1. DWG
Herramientas para la gestión y conversión de archivos CAD importados.
<br>
* **DWG Convert (En producción):** Convierte líneas y textos de DWG a los estándares de HMV.
<br>
* **3D DWG to Shape:** Extrae sólidos y mallas de un DWG 3D a un DirectShape nativo, se puede controlar el nivel de detalle.

### 2. Family Control Tools
Scripts para edición masiva y posicionamiento espacial de familias.
<br>
* **Ductos Editor:** Configurador visual interactivo de familia estandarizada de Banco de ductos.
<br>
* **Refresh Z to Floor (En producción):** Proyecta elementos hacia la cara superior de un suelo vinculado.
<br>
* **Multi InstParam Editor:** Edita parámetros de instancia de múltiples elementos en una sola transacción.
<br>
* **Foundation Control (En producción):** Ajusta la elevación de cimentaciones según niveles topográficos.

### 3. Annotation Tools
Automatización de tareas de documentación y etiquetado.
* **Grid/Level Extent:** Alterna masivamente entre extensión 2D y 3D en todas las vistas seleccionadas.
<br>
* **Topo to Lines:** Extrae curvas de nivel de un vínculo a líneas de detalle.
<br>
* **Align Spot Elevations:** Alinea cotas de elevación a un eje común.

### 4. Audit
Herramientas para mantener la salud y el estándar del modelo.
<br>
* **Text / Dim Audit:** Estandariza fuentes, tamaños y nombres de tipos.
<br>
* **View / Sheet Audit:** Renombra vistas y planos detectando duplicados.
<br>
* **Family Audit:** Compara versiones de familias contra la nube (ADC) usando hashes SHA-256.

---

> Usa el menú superior o la barra de búsqueda para ver el detalle y funcionamiento de cada comando específico.