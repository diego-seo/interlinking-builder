---
name: interlinking-builder
description: Genera estrategias de enlazado interno (interlinking) SEO a partir de listas de URLs en CSV o Excel.
---

# Interlinking Builder

Skill para generar planes de enlazado interno SEO desde cero a partir de una lista de URLs.

## Input esperado

| Columna | Descripcion | Requerida |
|---------|-------------|-----------|
| url | URL completa de la pagina | Si |
| titulo | Titulo o H1 | Recomendada |
| tipo | home, categoria, articulo, producto, landing | Opcional |
| keyword_principal | Keyword que ataca esa pagina | Opcional |
| volumen | Volumen de busqueda mensual | Opcional |

## Modelos de arquitectura

- **Hub-and-Spoke**: Paginas pilar + articulos de apoyo bidireccionales
- **Cadena**: Menor volumen hacia mayor volumen (webs pequenas)
- **SILO**: Categorias tematicas aisladas (e-commerce grande)
- **Vertizontal**: SILO + cadena interna (recomendado para >50 URLs)

## Output

Excel con 4 pestanas:
1. Plan_Interlinking: pares origen->destino con 3 anchor texts cada uno
2. Resumen_URLs: cobertura por URL
3. Alertas: huerfanas, saturacion, repetidos
4. Modelo_Aplicado: descripcion del modelo elegido

## Anchor texts (3 por par)

- Ancla 1: coincidencia exacta (5-10%)
- Ancla 2: semantica/LSI (25-35%)
- Ancla 3: long-tail descriptivo (55-70%)

Prohibido: "haz clic aqui", "leer mas", "aqui", "este articulo", etc.

Ver referencias:
- references/modelos-arquitectura.md
- references/anchor-text-avanzado.md
