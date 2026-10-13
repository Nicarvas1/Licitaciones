import importlib.util,sys,types,re,json
from pathlib import Path
sys.path.insert(0,".")
sp=importlib.util.spec_from_file_location("e","3_extraer_ia.py");e=importlib.util.module_from_spec(sp);sp.loader.exec_module(e)
A=types.SimpleNamespace(max_paginas=60,ocr=False,vision=False,tesseract_cmd=None,idioma_ocr="spa+eng",dpi_ocr=200,vision_min_filas=2,max_filas_excel=400,max_columnas_excel=25)
MONTO=re.compile(r"\d{1,3}(?:[.,]\d{3}){1,3}")
EQ=re.compile(r"notebook|laptop|port[aá]til|computador|desktop|all.?in.?one|\baio\b|workstation|monitor|impresora|multifuncional|pc\b|equipo|marca|modelo|thinkpad|probook|elitebook|macbook|zbook|latitude|ideapad|vivobook",re.I)
def hoja(base,lic,out):
    d=Path(base)/lic
    L=[f"##### LICITACION {lic}"]
    for prov in sorted(p for p in d.iterdir() if p.is_dir() and not p.name.startswith("_")):
        info=json.load(open(prov/"oferta.json",encoding="utf-8")) if (prov/"oferta.json").exists() else {}
        L.append(f"\n=== {prov.name} | total {info.get('total')} | estado {info.get('estado')}")
        for f in e.listar_archivos_oferta(prov):
            tipo=e.tipo_documental(f.name)
            try: t,est=e.extraer_archivo(f,A,{})
            except Exception as x: t,est="",f"error {x}"
            if not t.strip(): L.append(f"  [{tipo}] {f.name[:70]} -> SIN TEXTO ({est})"); continue
            L.append(f"  [{tipo}] {f.name[:70]} ({len(t)} chars)")
            lines=[l.strip() for l in t.splitlines() if l.strip()]
            if tipo=="economico":
                sel=[l for l in lines if MONTO.search(l) and (EQ.search(l) or l.startswith("FILA"))][:30]
            else:
                sel=[l for l in lines if EQ.search(l)][:8]
            for l in sel: L.append("      "+l[:230])
    open(out,"w",encoding="utf-8").write("\n".join(L))
if __name__=="__main__":
    base,out=sys.argv[1],sys.argv[2]
    for lic in sys.argv[3:]: hoja(base,lic,f"{out}/{lic}.txt")
