# Interlinking Builder

Skill de Diego González para preparar propuestas de enlazado interno a partir de un inventario de URLs. Incluye instrucciones para un asistente de IA y un script Python que genera un Excel con un plan inicial.

## Qué necesitas

- Un CSV separado por comas y guardado en UTF-8, o un archivo XLSX.
- Una columna `url` con URLs completas. Añade título, tipo y keyword para mejorar el contexto.
- Para ejecutar el script: Python 3, pandas y openpyxl. No requiere claves API.
- Para usar las instrucciones como skill: un asistente compatible. Sus límites y costes dependen del proveedor.

## Uso como skill en Claude Code

Descarga este repositorio desde **Code > Download ZIP**, descomprímelo y renombra la carpeta a `interlinking-builder`. Conserva todos sus archivos. Colócala dentro de `.claude/skills/` de tu proyecto. La ruta final debe ser:

```text
.claude/skills/interlinking-builder/SKILL.md
```

Para disponer del skill en todos tus proyectos locales, usa `~/.claude/skills/interlinking-builder/` (en Windows, dentro de tu carpeta de usuario). No sobrescribas otra versión sin guardar sus cambios.

Abre Claude Code en el proyecto e invócalo con `/interlinking-builder`, seguido de tu solicitud. Si no aparece, comprueba el nombre de la carpeta y que no haya una carpeta adicional entre ella y SKILL.md.

Esta instalación está documentada para Claude Code. Otros clientes compatibles con Agent Skills pueden requerir otra ubicación o un mecanismo de importación. No basta con pegar la URL de GitHub para instalarlo.

Referencia: [documentación oficial de skills de Claude Code](https://code.claude.com/docs/en/skills).

## Ejemplo de solicitud al asistente

> Usa interlinking-builder con mi archivo urls.csv. Propón enlaces relevantes para el lector, explica el modelo elegido y revisa manualmente los textos de anclaje. Distingue el plan propuesto de los enlaces existentes. No inventes volúmenes ni afirmes que una página está huérfana en la web sin un rastreo.

## Ejecutar el script sin asistente

Descarga y descomprime el repositorio o clónalo:

```sh
git clone https://github.com/diego-seo/interlinking-builder.git
cd interlinking-builder
python -m venv .venv
```

Activa el entorno en Windows PowerShell con `.\.venv\Scripts\Activate.ps1`; en macOS/Linux, con `source .venv/bin/activate`. Si PowerShell bloquea la activación, utiliza directamente `.\.venv\Scripts\python.exe` en lugar de `python`.

```sh
python -m pip install -r requirements.txt
python interlinking_builder.py --input ejemplo_urls.csv --output plan-ejemplo.xlsx
python interlinking_builder.py --help
```

En algunos sistemas el ejecutable se llama `python3`. Ejecuta los comandos desde la carpeta del repositorio. Elige una salida nueva: el script puede sobrescribir un Excel existente.

## Formato de entrada

| Columna | Uso |
| --- | --- |
| url | URL completa, una por fila. Obligatoria para un archivo fiable. |
| titulo | Título de la página, recomendado. |
| tipo | home, categoria, articulo, producto o landing. |
| keyword_principal | Término principal conocido de esa URL. |
| volumen | Entero sin separadores de miles; dejar vacío si no se conoce. |

Los datos de ejemplo son ficticios. Las URLs duplicadas se eliminan. Si faltan título, tipo o keyword, el script intenta inferirlos del slug; revisa esas inferencias.

## Modelos disponibles

Usa `--modelo auto`, `cadena`, `hub-and-spoke`, `silo` o `vertizontal`.

- Auto elige mediante reglas internas sobre cantidad y tipo de URLs. No es una recomendación SEO validada para tu negocio.
- Cadena ordena por volumen y utiliza coincidencias de palabras del slug.
- Hub-and-spoke relaciona pilares con artículos o productos por coincidencias del slug.
- En esta versión, silo ejecuta la misma función que hub-and-spoke; no garantiza aislamiento temático.
- Vertizontal agrupa por segmentos de la URL y crea cadenas dentro de cada grupo.

## Qué entrega

Un Excel con cuatro hojas: Plan_Interlinking, Resumen_URLs, Alertas y Modelo_Aplicado. Cada par propone tres textos de anclaje para revisar, no tres enlaces que debas insertar juntos.

## Limitaciones importantes

El script no visita las páginas, no conoce sus enlaces actuales, no comprueba códigos HTTP ni aplica cambios en tu CMS. La similitud utiliza palabras de las URLs; no analiza semánticamente el contenido completo.

Las etiquetas de huérfana, pocos enlaces y saturación solo describen el **plan generado**. No prueban problemas reales de indexación. El umbral de 44 enlaces y las proporciones de anchors incluidas en las instrucciones son reglas internas, no límites oficiales de Google. Tampoco se deben interpretar las referencias a LSI o a posiciones del enlace como garantías de posicionamiento.

Los anchors se construyen con plantillas y pueden ser repetitivos o poco naturales. Revisa relevancia, gramática y afirmaciones comerciales antes de utilizarlos. Un resultado sin alertas tampoco certifica la calidad SEO del plan.

Para esta versión usa CSV con comas o XLSX: la detección de otros separadores no es fiable. XLS antiguo requiere dependencias adicionales no incluidas. Revisa especialmente modelos con artículos que obtengan la misma similitud, pues esa combinación no está validada.

## Comprobación realizada

El 23 de septiembre de 2026 se ejecutó el ejemplo incluido con los cinco modos. Auto y cadena generaron 2 pares; hub-and-spoke y silo generaron 0; vertizontal generó 2. El ejemplo contiene solo 3 URLs y ningún artículo o producto de apoyo, por lo que no demuestra el comportamiento con inventarios grandes. Esta prueba comprueba ejecución, no resultados SEO.

## Privacidad y soporte

El script funciona con archivos locales y no contiene llamadas de red. Si compartes el inventario con un asistente de IA, aplican las condiciones de ese proveedor. Retira URLs privadas y datos de clientes que no deban compartirse.

Para informar un problema, abre un issue con el comando, el mensaje de error y un ejemplo anonimizado. No publiques credenciales ni datos privados.

## Licencia

[MIT](LICENSE). Puedes usar, modificar y redistribuir el recurso, incluso comercialmente, conservando el aviso de autoría y la licencia. Sin garantías de resultados.
