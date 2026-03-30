# Bienvenido 

**HMV Tools** es una suite de herramientas (Add-ins) personalizadas desarrolladas en C# para la API de Autodesk Revit. Su objetivo es automatizar tareas repetitivas, auditar modelos y estandarizar la información de los proyectos.

---

## Instalación
Para instalar las herramientas en tu equipo:

1. Descarga la última versión del instalador `.msi` desde el entorno, ruta del archivo debe ser asignada por el coordinador BIM.
<br><br>
[Ir a la carpeta en ACC](https://acc.autodesk.com/docs/files/projects/992dbb1c-6004-472d-a416-934fff1806d1?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.DX2kA7XYTIG91f0G8eY18g&viewModel=detail&moduleId=folders){ .md-button }
<br><br>
> **Nota:** El acceso a esta carpeta puede estar restringido.


2. Cierra Revit.

3. Ejecuta el instalador y sigue las instrucciones.

4. Al abrir Revit, encontrarás la nueva pestaña **HMV Tools** en la cinta de opciones (Ribbon).

---

## Resumen de Módulos

La cinta de opciones está dividida en 4 paneles principales para facilitar el flujo de trabajo.
<br>
En producción significa que la herramienta está en fase de desarrollo y aún no cumple su proposito en totalidad.

### 1. DWG
Herramientas para la gestión y conversión de archivos CAD importados (2D y 3D).
<br>
* **DWG Convert (En producción):** Convierte líneas , regiones de relleno y textos de DWG a los estándares de HMV.
<br>
* **3D DWG to Shape:** Extrae sólidos y mallas de un DWG 3D a un DirectShape nativo, se puede controlar el nivel de detalle.

### 2. Family Control Tools
Scripts para edición masiva ,posicionamiento y control espacial de familias.
<br>
* **Ductos Editor:** Configurador visual interactivo de familia estandarizada de Banco de ductos, también estandariza el nombre de tipo de esa familia.
<br>
* **Refresh Z to Floor (En producción):** Proyecta elementos hacia la cara superior de un suelo vinculado y actualiza el punto de intersección con el elemento.
<br>
* **Multi InstParam Editor:** Edita múltiples parámetros de instancia de elementos en una sola transacción, a fin de no perder elementos referenciados a estos(dimensionas, cotas, etiquetas).
<br>
* **Foundation Control (En producción):** Configurador visual interactivo para ajustar la elevación de cimentaciones según niveles NAP (Interseccion con el suelo de arquitectura) y NTCE (punto mas alto del elemento).

### 3. Annotation Tools
Automatización de tareas de documentación y etiquetado.

* **Topo to Lines:** Extrae topografía de un vínculo a líneas de detalle, genera el proceso de estandarizacion para adecuación a DMS.
<br>
* **Pipe Framing Annotations:** Permite replicar un Generic Annotation a traves de todo el recorrido de un Pipe, Flex Pipe o Framing.
<br>
* **Grid/Level Extent:** Permite cambiar de 2D a 3D la configuracion de un Grid en todas las vistas.
<br>
* **Spot Elev on Floor:** Calcula el punto de interseccion con el suelo (NAP) y punto mas alto de elemento (NTCE) y los deja acotados.
<br>
* **Generic Annot Tags:** Permite acotar un parametro que no sea compartido, lo inserta en forma de texto.
<br>
* **Align Spot Elevations:** Alinea cotas de elevación a un eje común.

### 4. Audit
Herramientas para mantener la salud y el estándar del modelo.
<br>
* **Text / Dim Audit:** Estandariza fuentes, tamaños y nombres de tipos.
<br>
* **View / Sheet Audit:** Renombra vistas y planos detectando duplicados.
<br>
* **Family Audit:** Compara versiones de familias contra la nube (ADC)´.

---

> Usa el menú superior o la barra de búsqueda para ver el detalle y funcionamiento de cada comando específico.