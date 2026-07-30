"""
INTERLINKING BUILDER - SEO Skill Tool
Genera planes de enlazado interno desde CSV o Excel con URLs.

USO:
    python interlinking_builder.py --input urls.csv
    python interlinking_builder.py --input urls.xlsx --modelo hub-and-spoke

MODELOS: auto, hub-and-spoke, cadena, silo, vertizontal
INSTALACION: pip install pandas openpyxl
"""

import argparse, os, sys, re
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("Falta pandas: pip install pandas"); sys.exit(1)

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Falta openpyxl: pip install openpyxl"); sys.exit(1)

COLORES = {
    "header_bg": "090909", "header_font": "BBF340",
    "alta_prioridad": "BBF340", "media_prioridad": "E9E9E9",
    "baja_prioridad": "F8F8F8", "alerta_roja": "FF4444",
    "alerta_amarilla": "FFD700", "texto_oscuro": "090909",
    "texto_gris": "666666", "borde": "D0D0D0",
}

COLUMN_MAP = {
    "url": ["url", "link", "pagina", "address", "href", "URL", "Page"],
    "titulo": ["titulo", "title", "h1", "nombre", "name", "Title"],
    "tipo": ["tipo", "type", "categoria", "category", "template"],
    "keyword": ["keyword", "keyword_principal", "kw", "query"],
    "volumen": ["volumen", "volume", "busquedas", "msv", "Volume"],
    "autoridad": ["autoridad", "authority", "da", "pa", "trafico", "traffic"],
}

ANCLAS_PROHIBIDAS = [
    "haz clic aqui", "clic aqui", "haga clic", "leer mas", "ver mas",
    "mas informacion", "aqui", "este articulo", "este post",
    "enlace", "link", "visita", "click here", "read more", "here",
]

MODELO_DESCRIPCIONES = {
    "hub-and-spoke": "Paginas pilar <-> articulos de apoyo (bidireccional). Ideal para blogs.",
    "cadena": "Enlace lineal menor volumen -> mayor volumen. Para webs pequenas.",
    "silo": "Categorias tematicas aisladas. Para e-commerce grande.",
    "vertizontal": "Hibrido SILO + cadena interna. Para >50 URLs multitematicas.",
}


def leer_archivo(filepath):
    ext = Path(filepath).suffix.lower()
    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(filepath, dtype=str)
    for enc in ["utf-8", "latin-1", "cp1252"]:
        for sep in [",", ";", "\t", "|"]:
            try:
                df = pd.read_csv(filepath, encoding=enc, sep=sep, dtype=str)
                if len(df.columns) >= 1:
                    return df
            except Exception:
                pass
    raise ValueError(f"No se pudo leer: {filepath}")


def normalizar_columnas(df):
    lower = {c.lower().strip(): c for c in df.columns}
    renames = {}
    for campo, variantes in COLUMN_MAP.items():
        for v in variantes:
            if v.lower() in lower:
                renames[lower[v.lower()]] = campo
                break
    df = df.rename(columns=renames)
    if "url" not in df.columns:
        df = df.rename(columns={df.columns[0]: "url"})
    for col in ["titulo", "tipo", "keyword", "volumen", "autoridad"]:
        if col not in df.columns:
            df[col] = ""
    df["url"] = df["url"].astype(str).str.strip()
    df = df[df["url"].notna() & (df["url"] != "") & (df["url"] != "nan")]
    return df.drop_duplicates(subset=["url"]).reset_index(drop=True)


def extraer_slug(url):
    slug = re.sub(r"https?://[^/]+", "", str(url))
    slug = re.sub(r"\.(html?|php|aspx?)$", "", slug)
    stop = {"como","que","para","una","los","las","del","con","por","en","de","la","el","y","a","the","and","of"}
    return [p.lower() for p in re.split(r"[-_/\s]+", slug) if p and p.lower() not in stop and len(p) > 2]


def inferir_titulo(row):
    if str(row.get("titulo", "")) not in ["", "nan"]:
        return str(row["titulo"]).strip()
    return " ".join(extraer_slug(row["url"])[:6]).title() or str(row["url"])


def inferir_tipo(row):
    t = str(row.get("tipo", ""))
    if t not in ["", "nan"]:
        return t.lower()
    url = str(row["url"]).lower()
    segs = [s for s in url.split("/") if s and s not in ["http:", "https:", ""]]
    if len(segs) <= 1: return "home"
    if any(p in url for p in ["/producto/", "/product/", "/shop/", "/tienda/"]): return "producto"
    if any(p in url for p in ["/servicio", "/service", "/landing", "/lp/"]): return "landing"
    if len(segs) <= 2: return "categoria"
    return "articulo"


def preparar(df):
    df["titulo_limpio"] = df.apply(inferir_titulo, axis=1)
    df["tipo_inferido"] = df.apply(inferir_tipo, axis=1)
    df["volumen_num"] = df.apply(lambda r: int(float(str(r.get("volumen","")).replace(".","").replace(",",""))) if str(r.get("volumen","")) not in ["","nan"] else 0, axis=1)
    df["slug_palabras"] = df["url"].apply(extraer_slug)
    df["kw"] = df.apply(lambda r: str(r["keyword"]).strip() if str(r.get("keyword","")) not in ["","nan"] else " ".join(r["slug_palabras"][:4]), axis=1)
    return df


def elegir_modelo(df, forzado="auto"):
    if forzado != "auto": return forzado
    n, tipos = len(df), df["tipo_inferido"].value_counts().to_dict()
    n_tipos = len([t for t,c in tipos.items() if c > 0 and t != "home"])
    if n <= 20: return "cadena"
    if n <= 50 and n_tipos <= 2: return "hub-and-spoke"
    if tipos.get("categoria",0) > 0 and n_tipos >= 3: return "vertizontal" if n > 50 else "silo"
    return "hub-and-spoke"


def simil(a, b):
    sa, sb = set(a), set(b)
    return len(sa & sb) / len(sa | sb) if (sa | sb) else 0.0


def generar_anclas(row):
    tipo, kw, titulo = row["tipo_inferido"], row["kw"], row["titulo_limpio"]
    extra = [p for p in titulo.split() if p.lower() not in kw.lower() and len(p) > 3]
    a1 = kw or titulo[:40]
    a2 = (f"{kw} {extra[0]}" if extra else f"guia de {kw}").strip()
    if tipo == "landing": a2 = f"servicios de {kw}"
    elif tipo == "categoria": a2 = f"todo sobre {kw}"
    if tipo == "articulo": a3 = f"como {kw} paso a paso"
    elif tipo == "producto": a3 = f"comprar {titulo} al mejor precio"
    elif tipo == "landing": a3 = f"contratar {kw} para tu negocio"
    elif tipo == "categoria": a3 = f"ver todos los recursos sobre {kw}"
    else: a3 = f"aprende mas sobre {kw} en esta guia"
    return a1[:80], a2[:80], a3[:120]


def pares_hub_spoke(df):
    pares = []
    pilares = df[df["tipo_inferido"].isin(["categoria","landing","home"])]
    arts = df[df["tipo_inferido"].isin(["articulo","producto"])]
    if pilares.empty:
        pilares = df.nlargest(max(1,len(df)//5), "volumen_num")
        arts = df[~df["url"].isin(pilares["url"])]
    for _, p in pilares.iterrows():
        rel = sorted([(simil(p["slug_palabras"],a["slug_palabras"]),a) for _,a in arts.iterrows() if simil(p["slug_palabras"],a["slug_palabras"])>0.1], reverse=True)[:6]
        for sc, a in rel:
            pr = "Alta" if sc > 0.3 else "Media"
            pares.append({"url_origen":a["url"],"titulo_origen":a["titulo_limpio"],"url_destino":p["url"],"titulo_destino":p["titulo_limpio"],"tipo_enlace":"Apoyo->Pilar","prioridad":pr,"donde_insertar":"Cuerpo"})
            pares.append({"url_origen":p["url"],"titulo_origen":p["titulo_limpio"],"url_destino":a["url"],"titulo_destino":a["titulo_limpio"],"tipo_enlace":"Pilar->Apoyo","prioridad":"Media","donde_insertar":"Intro"})
    return pares


def pares_cadena(df):
    pares = []
    s = df.sort_values("volumen_num").reset_index(drop=True)
    for i in range(len(s)-1):
        o = s.iloc[i]; d = s.iloc[i+1]
        if simil(o["slug_palabras"],d["slug_palabras"]) < 0.1:
            best = max(range(i+2,min(i+5,len(s))), key=lambda j: simil(o["slug_palabras"],s.iloc[j]["slug_palabras"]), default=i+1)
            d = s.iloc[best]
        pares.append({"url_origen":o["url"],"titulo_origen":o["titulo_limpio"],"url_destino":d["url"],"titulo_destino":d["titulo_limpio"],"tipo_enlace":"Cadena (vol)","prioridad":"Alta" if d["volumen_num"]>500 else "Media","donde_insertar":"Intro (25%)"})
    return pares


def pares_vertizontal(df):
    pares = []
    clusters = {}
    for _, r in df.iterrows():
        segs = [s for s in r["url"].replace("https://","").replace("http://","").split("/") if s and "." not in s]
        c = segs[0] if len(segs)>1 else "general"
        clusters.setdefault(c,[]).append(r["url"])
    for cn, urls in clusters.items():
        sub = df[df["url"].isin(urls)].copy()
        if len(sub) < 2: continue
        for p in pares_cadena(sub):
            p["tipo_enlace"] = f"Vertizontal [{cn}]"
            pares.append(p)
        home = df[df["tipo_inferido"]=="home"]
        if not home.empty:
            rep = sub.nlargest(1,"volumen_num").iloc[0]
            h = home.iloc[0]
            pares.append({"url_origen":rep["url"],"titulo_origen":rep["titulo_limpio"],"url_destino":h["url"],"titulo_destino":h["titulo_limpio"],"tipo_enlace":"Cluster->Home","prioridad":"Media","donde_insertar":"Conclusion"})
    return pares


def generar_pares(df, modelo):
    if modelo == "cadena": ps = pares_cadena(df)
    elif modelo in ["silo","hub-and-spoke"]: ps = pares_hub_spoke(df)
    else: ps = pares_vertizontal(df)
    url_row = {r["url"]: r for _, r in df.iterrows()}
    for p in ps:
        if p["url_destino"] in url_row:
            p["ancla_1"],p["ancla_2"],p["ancla_3"] = generar_anclas(url_row[p["url_destino"]])
        else:
            p["ancla_1"] = p["ancla_2"] = p["titulo_destino"]
            p["ancla_3"] = f"mas sobre {p['titulo_destino']}"
    return ps


def cobertura(df, pares):
    ent,sal = {},{}
    for p in pares:
        sal[p["url_origen"]] = sal.get(p["url_origen"],0)+1
        ent[p["url_destino"]] = ent.get(p["url_destino"],0)+1
    rows = []
    for _,r in df.iterrows():
        e,s = ent.get(r["url"],0),sal.get(r["url"],0)
        t = e+s
        rows.append({"url":r["url"],"titulo":r["titulo_limpio"],"tipo":r["tipo_inferido"],
                     "entrantes":e,"salientes":s,"total":t,
                     "estado":"🚨 Huerfana" if e==0 else "⚠️ Pocos" if t<3 else "✅ OK"})
    return pd.DataFrame(rows)


def alertas(df_cob, pares):
    al = []
    for _,r in df_cob[df_cob["estado"]=="🚨 Huerfana"].iterrows():
        al.append({"tipo":"🚨 Huerfana","url":r["url"],"detalle":"Sin enlaces entrantes. Riesgo de no indexarse.","accion":"Enlazar desde articulo del mismo tema."})
    for _,r in df_cob[df_cob["total"]>44].iterrows():
        al.append({"tipo":"⚠️ Saturacion","url":r["url"],"detalle":f"Tiene {r['total']} enlaces (max recomendado: 44).","accion":"Reducir enlaces menos relevantes."})
    anclas = {}
    for p in pares:
        anclas.setdefault(p["url_destino"],[]).append(p.get("ancla_1",""))
    for dest,a in anclas.items():
        if len(a) != len(set(a)):
            al.append({"tipo":"⚠️ Anchor repetido","url":dest,"detalle":"Misma Ancla 1 en varios enlaces a esta URL.","accion":"Diversificar con Ancla 2 o Ancla 3."})
    return al


def header_style(ws, row=1):
    for c in ws[row]:
        c.fill = PatternFill("solid",fgColor=COLORES["header_bg"])
        c.font = Font(bold=True,color=COLORES["header_font"],size=10)
        c.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)


def exportar(pares, df_cob, als, modelo, path):
    wb = Workbook()
    # Sheet 1
    ws1 = wb.active; ws1.title = "Plan_Interlinking"
    ws1.append(["#","URL Origen","Titulo Origen","URL Destino","Titulo Destino","Ancla 1 Exacta","Ancla 2 Semantica","Ancla 3 Long-tail","Tipo Enlace","Prioridad","Donde Insertar"])
    header_style(ws1)
    cp = {"Alta":COLORES["alta_prioridad"],"Media":COLORES["media_prioridad"],"Baja":COLORES["baja_prioridad"]}
    for i,p in enumerate(pares,1):
        pr = p.get("prioridad","Media")
        ws1.append([i,p.get("url_origen",""),p.get("titulo_origen",""),p.get("url_destino",""),p.get("titulo_destino",""),p.get("ancla_1",""),p.get("ancla_2",""),p.get("ancla_3",""),p.get("tipo_enlace",""),pr,p.get("donde_insertar","")])
        fill = PatternFill("solid",fgColor=cp.get(pr,COLORES["baja_prioridad"]))
        for c in ws1[ws1.max_row]: c.fill=fill; c.font=Font(color=COLORES["texto_oscuro"],size=9)
    ws1.freeze_panes="A2"
    # Sheet 2
    ws2 = wb.create_sheet("Resumen_URLs")
    ws2.append(["URL","Titulo","Tipo","Entrantes","Salientes","Total","Estado"]); header_style(ws2)
    for _,r in df_cob.iterrows():
        ws2.append([r["url"],r["titulo"],r["tipo"],r["entrantes"],r["salientes"],r["total"],r["estado"]])
        fc = COLORES["alerta_roja"] if "Huerfana" in r["estado"] else COLORES["alerta_amarilla"] if "Pocos" in r["estado"] else COLORES["baja_prioridad"]
        for c in ws2[ws2.max_row]: c.fill=PatternFill("solid",fgColor=fc); c.font=Font(color=COLORES["texto_oscuro"],size=9)
    ws2.freeze_panes="A2"
    # Sheet 3
    ws3 = wb.create_sheet("Alertas")
    ws3.append(["Tipo","URL","Detalle","Accion"]); header_style(ws3)
    if als:
        for a in als:
            ws3.append([a["tipo"],a["url"],a["detalle"],a["accion"]])
            fc = COLORES["alerta_roja"] if "Huerfana" in a["tipo"] else COLORES["alerta_amarilla"]
            for c in ws3[ws3.max_row]: c.fill=PatternFill("solid",fgColor=fc); c.font=Font(color=COLORES["texto_oscuro"],size=9)
    else:
        ws3.append(["✅ Sin alertas","","Plan bien balanceado.",""])
    # Sheet 4
    ws4 = wb.create_sheet("Modelo_Aplicado")
    ws4.column_dimensions["A"].width=80
    t=ws4.cell(1,1,f"MODELO: {modelo.upper()}"); t.fill=PatternFill("solid",fgColor=COLORES["header_bg"]); t.font=Font(bold=True,color=COLORES["header_font"],size=12)
    ws4.cell(2,1,MODELO_DESCRIPCIONES.get(modelo,"")).font=Font(size=10)
    ws4.cell(4,1,"RESUMEN").font=Font(bold=True,size=10)
    for i,(l,v) in enumerate([(f"URLs procesadas",len(df_cob)),(f"Pares de enlace",len(pares)),(f"URLs huerfanas",len(df_cob[df_cob["estado"]=="🚨 Huerfana"])),(f"Alertas totales",len(als))],5):
        ws4.cell(i,1,f"  {l}: {v}").font=Font(size=10)
    wb.save(path)


def main():
    p = argparse.ArgumentParser(description="Interlinking Builder - Plan de enlazado interno SEO")
    p.add_argument("--input","-i",required=True)
    p.add_argument("--output","-o",default=None)
    p.add_argument("--modelo","-m",default="auto",choices=["auto","hub-and-spoke","cadena","silo","vertizontal"])
    args = p.parse_args()
    if not os.path.exists(args.input): print(f"Archivo no encontrado: {args.input}"); sys.exit(1)
    out = args.output or f"{Path(args.input).stem}_interlinking.xlsx"
    print(f"\nINTERLINKING BUILDER")
    print(f"Input: {args.input} | Output: {out} | Modelo: {args.modelo}\n")
    df = leer_archivo(args.input)
    df = normalizar_columnas(df)
    print(f"URLs detectadas: {len(df)}")
    df = preparar(df)
    modelo = elegir_modelo(df, args.modelo)
    print(f"Modelo: {modelo.upper()} - {MODELO_DESCRIPCIONES[modelo]}")
    pares = generar_pares(df, modelo)
    print(f"Pares generados: {len(pares)}")
    df_cob = cobertura(df, pares)
    als = alertas(df_cob, pares)
    exportar(pares, df_cob, als, modelo, out)
    huerfanas = len(df_cob[df_cob["estado"]=="🚨 Huerfana"])
    print(f"\n✅ Listo: {out}")
    print(f"  • {len(df)} URLs | {len(pares)} pares | {len(als)} alertas | {huerfanas} huerfanas\n")


if __name__ == "__main__":
    main()
