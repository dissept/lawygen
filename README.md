# LawyGen — Generador de Excel, Resumen y Esquema legal

Aplicación de escritorio que analiza carpetas de casos legales y genera,
de forma independiente con tres botones:

1. **📊 Generar Excel** — un Excel maestro (`Seguimiento_Legal.xlsx`) con una
   fila por procedimiento: Asunto, Procedimiento, **Materia** (civil/penal/
   laboral/mercantil/contencioso-admvo.), Juzgado, Partes, **Representantes**
   (abogado/procurador de cada parte), **Argumentos de las partes**, **Objeto
   / causa juzgada**, **Fundamentos de Derecho**, **Cuantía y costas**, columna **Estado**
   con desplegable real, hipervínculo directo al PDF principal de cada
   carpeta y una columna con los **demás documentos/adjuntos** de esa carpeta.
2. **📝 Generar Resumen** — **un único** Word `Resumen_General.docx` guardado
   en la carpeta raíz, con una sección por cada asunto: tabla de datos
   completa (incluye materia, representantes, objeto, fundamentos de derecho,
   cuantía y costas), un **resumen ejecutivo** en viñetas y una sección de
   **"Argumentos de las partes"** cuando se detectan.
3. **🗂️ Generar Esquema** — un Word `<Carpeta>_Esquema.docx` por cada asunto,
   con un **diagrama de flujo visual** (imagen) de las fases y cláusulas
   detectadas en el procedimiento.

Por defecto **todo el procesamiento ocurre en el propio ordenador** (100%
gratis y offline, sin IA de pago): ningún documento sale del equipo. Ver la
sección 5 para activar de forma opcional un análisis asistido por IA que
mejora la precisión de estos campos.

---

## 1. Instalación (una sola vez, por cada PC del despacho)

Requisitos: [Python 3.10 o superior](https://www.python.org/downloads/)
(al instalar, marcar la casilla **"Add Python to PATH"**).

Abre una terminal (`cmd` en Windows) dentro de esta carpeta y ejecuta:

```
pip install -r requirements.txt
```

## 2. Uso

```
python app.py
```

Se abrirá una ventana morada de **LawyGen**:

1. Clic en **"Elegir carpeta..."** y selecciona la carpeta que contiene las
   subcarpetas de cada cliente/asunto (una subcarpeta = un asunto).
2. (Opcional) En **"2. IA (opcional)"**, pega tu propia clave de la API de
   Anthropic y pulsa **"Guardar clave"** si quieres mejorar la precisión —
   ver sección 5. Si no la rellenas, la app sigue funcionando gratis.
3. Pulsa el botón correspondiente a lo que necesites generar:
   - **📊 Generar Excel** → crea `Seguimiento_Legal.xlsx` en la carpeta raíz.
   - **📝 Generar Resumen** → crea **un solo** `Resumen_General.docx` en la
     carpeta raíz, con todos los asuntos organizados dentro.
   - **🗂️ Generar Esquema** → crea `<Carpeta>_Esquema.docx` (con diagrama
     visual) dentro de cada subcarpeta.

Los tres botones son independientes: puedes pulsar solo el que necesites, o
los tres, en cualquier orden.

## 3. Generar el .exe SIN tener un PC con Windows (recomendado si trabajas desde Linux/Kali)

Este proyecto incluye `.github/workflows/build-exe.yml`, que compila el `.exe`
automáticamente en un Windows en la nube de GitHub (gratis). Pasos:

1. Crea una cuenta en [github.com](https://github.com) si no tienes una (gratis).
2. Crea un repositorio nuevo (puede ser **privado**, para que nadie más vea
   el código): botón "New repository".
3. Desde tu Kali, dentro de la carpeta `legal-doc-processor`, ejecuta:
   ```bash
   git init
   git add .
   git commit -m "Primera version"
   git branch -M main
   git remote add origin https://github.com/TU-USUARIO/TU-REPO.git
   git push -u origin main
   ```
4. Entra a tu repositorio en GitHub → pestaña **"Actions"**. Verás que el
   workflow "Build Windows EXE" se ejecuta solo (tarda 1-2 minutos).
5. Cuando termine (icono verde ✔), entra al resultado (el run que aparece) y
   baja hasta **"Artifacts"** → descarga `LawyGen-Windows.zip`.
6. Descomprímelo: ahí está `LawyGen.exe`, ya compilado para Windows.
7. Reparte ese único archivo a los PCs del departamento legal — no necesitan
   instalar Python ni nada más, doble clic y funciona.

Cada vez que modifiques el código y hagas `git push`, se genera un `.exe`
actualizado automáticamente.

### Alternativa: Wine en Kali (sin GitHub)

Si prefieres no usar GitHub, puedes compilar el `.exe` localmente con Wine:

```bash
sudo apt install wine
wine --version   # verifica que funciona
# Descarga el instalador de Python para Windows (python.org, versión .exe)
wine python-3.11.x-amd64.exe   # instálalo dentro de Wine, siguiendo el asistente
wine python -m pip install -r requirements.txt pyinstaller
wine pyinstaller --noconfirm --onefile --windowed --name "SeguimientoLegal" app.py
```
El `.exe` resultante queda en `dist/LawyGen.exe`. Este método es más
manual y depende de que Wine emule correctamente Python+Tkinter; si da
problemas, la opción de GitHub Actions es más fiable.

## 4. Cómo funciona la extracción por defecto (reglas de texto)

La app usa **patrones de texto (expresiones regulares) y reglas propias del
lenguaje jurídico español**, no inteligencia artificial, a menos que actives
la IA (ver sección 5):

- **Juzgado**: busca patrones como "Juzgado de...", "Audiencia Provincial
  de...", "Sala de lo... de...".
- **Partes**: busca etiquetas como "Demandante:", "Demandado:", "Actor:",
  o el patrón "X contra Y".
- **Representantes**: busca etiquetas "Abogado:", "Letrado:", "Procurador:"
  junto a cada parte.
- **Materia**: clasifica el asunto como Civil, Penal, Laboral, Mercantil o
  Contencioso-Administrativo según palabras clave típicas de cada
  jurisdicción.
- **Procedimiento**: busca "Procedimiento:", "Juicio Ordinario/Verbal...",
  "Diligencias Previas...", números de autos/expediente.
- **Objeto**: busca encabezados "Objeto de la demanda", "Suplico",
  "Peticiones" y extrae el fragmento inmediatamente posterior.
- **Fundamentos de Derecho**: busca encabezados "Fundamentos de Derecho",
  "Alegaciones", "Excepción Procesal", "Otrosí Digo" y extrae los primeros
  puntos como viñetas.
- **Cuantía y costas**: busca "Cuantía:", "por la cantidad de...", "con/sin
  imposición de costas".
- **Estado**: sugiere un valor según palabras clave ("sentencia firme" →
  Resuelto, "se archiva" → Archivado, etc.), pero **siempre revisable** con
  el desplegable de Excel.
- **Resumen / Argumentos**: extractivo, selecciona las frases más
  representativas del documento (sin generar texto nuevo, solo selecciona
  lo más relevante). Sin IA, la columna "Argumentos de las partes" queda en
  "revisar" (requiere IA para separar los argumentos por cada parte).
- **Esquema**: detecta encabezados típicos (Hechos, Fundamentos de Derecho,
  Suplico, Fallo, cláusulas "Primero.-", "Segundo.-", etc.).
- **Documentos adjuntos**: cada carpeta puede tener varios documentos; el
  Excel enlaza el principal (el PDF más grande) y lista los demás en la
  columna "Otros documentos adjuntos".

**Cuando un dato no se puede determinar con certeza, se marca como
"revisar"** en vez de inventarlo.

### Límites a tener en cuenta

- Si un PDF es un **escaneo sin texto** (imagen pura, sin OCR), no se podrá
  extraer texto y esos campos quedarán en "revisar". Si esto pasa con
  frecuencia, se puede añadir reconocimiento óptico (OCR) más adelante.
- El documento "principal" de cada carpeta se elige por tamaño de archivo
  (el PDF más grande). Si esa heurística falla en algún caso, se puede
  ajustar manualmente la celda de hipervínculo en el Excel.
- Al ser reglas de texto (cuando la IA está desactivada), funciona mejor
  cuanto más estandarizado sea el formato de los escritos/sentencias.

## 5. Activar la IA (opcional, mejora la precisión)

LawyGen incluye ya integrada la llamada a la API de Claude (Anthropic) para
mejorar la extracción de: partes/representantes, materia, objeto,
fundamentos de derecho, cuantía/costas, argumentos de cada parte, el
resumen ejecutivo y los documentos de las subcarpetas "Demandado"/
"Demandante" si existen. **Está
desactivada por defecto** y solo se activa si tú añades tu propia clave:

1. Consigue una clave de API en [console.anthropic.com](https://console.anthropic.com/)
   (requiere una cuenta de Anthropic con saldo; el coste es de fracciones
   de centavo por documento con el modelo por defecto).
2. Abre LawyGen, pega la clave en el campo **"2. IA (opcional)"** y pulsa
   **"Guardar clave"**. La clave se guarda solo en tu equipo (no se sube a
   ningún repositorio ni se comparte).
3. Deja marcada la casilla **"Usar IA cuando haya clave guardada"**. A
   partir de ahora, cada vez que proceses una carpeta, LawyGen enviará el
   texto de sus documentos a la API de Anthropic para completar los campos
   que las reglas de texto no puedan determinar con seguridad.
4. Si quieres volver al modo 100% gratuito, desmarca esa casilla o borra la
   clave guardada.

Si la IA falla por cualquier motivo (sin conexión, clave inválida, límite
de uso alcanzado...), LawyGen **cae automáticamente de vuelta a las reglas
de texto** para esa carpeta y sigue funcionando con normalidad; el aviso
queda anotado en el "Registro de actividad".

La clave se guarda localmente en:
- Windows: `%APPDATA%\LawyGen\config.json`
- Linux/Mac: `~/.LawyGen/config.json`

También puedes definir la variable de entorno `ANTHROPIC_API_KEY` en vez de
usar el campo de la app, si lo prefieres para un despliegue centralizado.

### Coste: Sonnet 5 vs Haiku 4.5

`ai_config.py` viene configurado por defecto con **Claude Sonnet 5**,
priorizando la máxima calidad de extracción (para que no tengas que revisar
casi nada a mano) sobre el coste o la velocidad. Envía más contexto de cada
documento y deja el razonamiento del modelo activado, así que en carpetas
grandes puede tardar más y costar algo más por carpeta — sigue siendo del
orden de céntimos, no dólares. Si en algún momento quieres priorizar el
coste sobre la precisión para tandas muy grandes, cambia `AI_MODEL` en
`ai_config.py` a `"claude-haiku-4-5-20251001"` (aprox. la mitad de precio).

## Estructura de archivos

```
app.py              → Interfaz gráfica (punto de entrada), incluye el panel de Ajustes de IA
processor.py         → Orquesta el análisis de cada carpeta y combina reglas + IA
legal_parser.py       → Reglas de extracción (Juzgado, Partes, Representantes, Materia,
                         Objeto, Fundamentos de Derecho, Cuantía/Costas,
                         Resumen, Esquema)
ai_extractor.py        → Llamada opcional a la API de Anthropic (Claude) y fusión con las reglas
config_store.py         → Guarda/lee localmente la clave de API y los ajustes de IA
extractor.py          → Lectura de texto de PDF/DOCX/TXT
excel_builder.py       → Generación del Excel con validación de datos e hipervínculos
build_exe.bat          → Empaqueta la app como .exe standalone (Windows)
requirements.txt        → Dependencias de Python
```

