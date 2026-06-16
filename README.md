<h1 align="center">🤖 Robótica</h1>

<p align="center">
  <b>Generación Automatizada de Presentaciones Técnicas y Apuntes Académicos de Alta Gama</b>
  <br>
  <i>Desarrollado para la materia <b>Tecnologías para la Automatización</b> (4to Año, Ingeniería en Sistemas de Información).</i>
</p>

<hr>

## 🚀 Descripción del Proyecto

Este repositorio aloja herramientas automatizadas en Python diseñadas para generar de forma dinámica y estructurada material pedagógico de alta fidelidad técnica. Utiliza el contenido del temario oficial de robótica industrial para construir automáticamente:

1. **Presentaciones Corporativas (`.pptx`)**: Diapositivas en formato panorámico (16:9) estructuradas jerárquicamente para exposiciones académicas y profesionales.
2. **Apuntes y Guías de Estudio (`.docx`)**: Documentos detallados con formato tipográfico unificado utilizando fuentes premium para su lectura y distribución.

---

## 📂 Contenido del Repositorio

El proyecto consta de los siguientes archivos principales:

*   **[`armar_pwp.py`](file:///c:/Users/Agus/Documents/UTN/Robotica/armar_pwp.py)**: Script encargado de generar la presentación de PowerPoint (`Robotica_Industrial_Completo.pptx`) que abarca exactamente las 99 secciones temáticas del curso. Cuenta con formateo de color de texto, numeración dinámica y etiquetas en color rojo para identificar dónde insertar recursos visuales.
*   **[`armar_docx.py`](file:///c:/Users/Agus/Documents/UTN/Robotica/armar_docx.py)**: Script encargado de ensamblar el apunte teórico unificado (`Robotica_Apunte_GoogleSans.docx`), estructurado mediante jerarquías de títulos, estilos normalizados y configuración nativa de la tipografía **Google Sans**.
*   **`README.md`**: Guía y documentación del repositorio.

---

## 🎓 Programa Temático Cubierto (99 Secciones)

El contenido generado por los scripts cubre la totalidad del espectro de robótica e integración industrial:

1.  **Fundamentos e Integración Industrial**
    *   Procesos productivos, automatización y objetivos estratégicos.
    *   Definición de robot, atributos (movilidad, versatilidad, percepción, autonomía).
    *   Leyes de la robótica de Isaac Asimov (1944).
    *   Espacio de trabajo, grados de libertad (GDL) y anatomía (eslabones y articulaciones).
    *   Efectores finales (grippers, herramientas, pistolas de rocío, sensores de medición).
2.  **Aplicaciones Industriales Comunes**
    *   Pintura automotriz, alimentación de máquinas, fundición, soldadura, montaje y paletización.
    *   Clasificación geométrica de manipuladores (Cartesianos, Cilíndricos, Esféricos, SCARA, Antropomórficos).
3.  **Localización y Cinemática Espacial**
    *   Posición, orientación y sistemas de referencia.
    *   Regla de la mano derecha y transformaciones espaciales en coordenadas cilíndricas y esféricas.
    *   Deducción de matrices de rotación en 2D y 3D (Ejes X, Y, Z).
    *   Matriz de Transformación Homogénea (4x4) y composición de movimientos.
4.  **Modelado Cinemático y Algoritmo de Denavit-Hartenberg**
    *   Fundamentos del Algoritmo de Denavit-Hartenberg (D-H, 1955).
    *   Estudio de los 4 parámetros fundamentales: $\theta$ (rotación de articulación), $d$ (distancia de enlace), $a$ (longitud de enlace) y $\alpha$ (torsión de enlace).
    *   Reglas y pasos para la asignación sistemática de ejes coordenados.
    *   Resolución analítica y matricial de configuraciones: Robot Planar (2 GDL) y Robot Cilíndrico (3 GDL).
5.  **Cinemática Inversa y Diferencial**
    *   Problema Cinemático Inverso (PCI), no linealidad y multiplicidad de soluciones.
    *   Técnica de desacoplo cinemático mediante muñeca esférica.
    *   Modelo diferencial: obtención y análisis de la Matriz Jacobiana.
    *   Identificación de singularidades jacobianas en el espacio de trabajo.
6.  **Dinámica, Planificación y Control**
    *   Ecuación dinámica general de manipuladores (matriz de inercia, términos centrífugos, Coriolis y fuerzas de gravedad).
    *   Formulaciones dinámicas: Lagrange-Euler vs. Newton-Euler.
    *   Control cinemático y perfiles de trayectoria (Punto a Punto - PTP, Trayectorias Continuas - CP).
    *   Interpoladores polinómicos (cúbicos) y perfiles de velocidad trapezoidales.
    *   Sistemas de lazo cerrado: Control PID y Control por Par Calculado.
7.  **Componentes de Hardware y Simulación**
    *   Sensores propioceptivos (encoders ópticos absolutos/incrementales, resolvers) y exteroceptivos (sensores de fuerza/torque, visión artificial).
    *   Transmisiones mecánicas de precisión y cero holgura (*Harmonic Drive*).
    *   Actuadores industriales (servomotores Brushless AC).
    *   Programación de robots: guiado (Teach Pendant) y textual (programación offline).
    *   Entornos modernos de simulación virtual y CAD (ej. ABB RobotStudio).
8.  **Criterios de Diseño, Seguridad y Negocios**
    *   Definición y contraste de precisión vs. repetibilidad.
    *   Optimización de celdas robotizadas, layouts de manufactura y tiempos de ciclo.
    *   Normas internacionales de seguridad industrial en celdas de trabajo.
    *   Robots colaborativos modernos (Cobots) e interacción humano-robot.
    *   Evaluación económica de proyectos de automatización (cálculo de ROI).
    *   Tendencias globales de Industria 4.0: computación en el borde (Edge) e Inteligencia Artificial en manufactura.

---

## 🛠️ Requisitos de Instalación

El proyecto está programado en **Python 3.8+** y requiere las siguientes librerías de generación de documentos de Office:

```bash
pip install python-pptx python-docx
```

### ⚠️ Nota sobre Fuentes y Visualización (Google Sans)

El script de generación de documentos (`armar_docx.py`) utiliza de forma nativa la tipografía **Google Sans** para todos los estilos de título y texto normal.
*   **Generación**: El script se ejecutará correctamente en cualquier sistema sin requerir descargas adicionales.
*   **Visualización**: Para poder ver correctamente el documento formateado con esta tipografía al abrir el archivo `.docx` generado en Microsoft Word o LibreOffice, **debes tener instalada la fuente Google Sans en tu sistema operativo**. De lo contrario, el software de procesamiento de textos utilizará fuentes alternativas por defecto (como *Calibri* o *Arial*).

---

## 🚀 Guía de Uso

Clona el repositorio e ingresa a la carpeta del proyecto:

```bash
git clone https://github.com/agussantinelli/Robotica.git
cd Robotica
```

### 1. Generar Presentación de Diapositivas (.pptx)

Ejecuta el script de generación de PowerPoint:

```bash
python armar_pwp.py
```

*   **Resultado**: Se creará un archivo llamado `Robotica_Industrial_Completo.pptx` en la carpeta raíz.
*   **Formato**: Contiene las 99 diapositivas estructuradas con numeración y etiquetas rojas donde se sugiere insertar imágenes o gráficos específicos.

### 2. Generar Apunte de Estudio (.docx)

Ejecuta el script de generación de documentos de Word:

```bash
python armar_docx.py
```

*   **Resultado**: Se creará un archivo llamado `Robotica_Apunte_GoogleSans.docx` en la carpeta raíz.
*   **Formato**: Un manual técnico unificado de 99 secciones enlazadas, configurado con títulos azules de contraste y tipografía corporativa.

---

<p align="center">
  <b>UTN - Facultad Regional - Departamento de Sistemas</b><br>
  📚 <i>Tecnología, matemáticas y automatización aplicadas al diseño curricular de vanguardia.</i>
</p>
