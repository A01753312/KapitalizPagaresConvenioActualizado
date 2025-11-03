# =========================================================================================
# APP.PY FINAL — UNA SOLA PLANTILLA DE PAGARÉ + DOMICILIO POR EXCEL
# =========================================================================================

import io, os, re, zipfile, tempfile, subprocess, shutil
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# ========== Utilidades originales ==========
SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9._\\-áéíóúÁÉÍÓÚñÑ ]+")

def safe_name(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = SAFE_NAME_RE.sub("_", s)
    s = re.sub(r"\s+", " ", s)
    return s[:150]

@st.cache_data(show_spinner=False)
def read_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    if sum(str(c).startswith("Unnamed") for c in df.columns) > len(df.columns) * 0.6:
        headers = [str(x).strip() for x in df.iloc[0].tolist()]
        df = df.iloc[1:].copy()
        df.columns = headers
    df.columns = [str(c).strip() for c in df.columns]
    return df

def normalize_str(x: str) -> str:
    if x is None:
        return ""
    s = str(x).lower().strip()
    for a, b in (("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ñ","n")):
        s = s.replace(a, b)
    return s

# Fecha “DD DE MES DEL YYYY”
MESES_MAYUS = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]

def fecha_hoy_es() -> str:
    d = datetime.now()
    return f"{d.day} DE {MESES_MAYUS[d.month-1]} DEL {d.year}"

# -------- Conversión de números a letras ----------
UNIDADES = ["cero","uno","dos","tres","cuatro","cinco","seis","siete","ocho","nueve","diez","once","doce","trece","catorce","quince","dieciséis","diecisiete","dieciocho","diecinueve","veinte"]
DECENAS  = ["","","veinte","treinta","cuarenta","cincuenta","sesenta","setenta","ochenta","noventa"]
CENTENAS = ["","cien","doscientos","trescientos","cuatrocientos","quinientos","seiscientos","setecientos","ochocientos","novecientos"]

def _tens(n:int)->str:
    if n<=20: return UNIDADES[n]
    d,u = divmod(n,10)
    if u==0: return DECENAS[d]
    if d==2: return f"veinti{UNIDADES[u]}".replace("veintiuno","veintiún")
    return f"{DECENAS[d]} y {UNIDADES[u]}"

def _hundreds(n:int)->str:
    if n==0: return ""
    if n==100: return "cien"
    c,r = divmod(n,100)
    pref = "ciento" if c==1 else CENTENAS[c]
    return (pref + (f" {_tens(r)}" if r else "")).strip()

def numero_a_letras(n:int)->str:
    if n==0: return "cero"
    partes=[]
    millones, r = divmod(n, 1_000_000)
    miles, unidades = divmod(r, 1000)
    if millones: partes.append("un millón" if millones==1 else f"{_hundreds(millones)} millones")
    if miles:    partes.append("mil" if miles==1 else f"{_hundreds(miles)} mil")
    if unidades: partes.append(_hundreds(unidades))
    return " ".join(partes).replace("uno mil","un mil").replace("veintiun","veintiún")

def monto_en_letras(mx: float)->str:
    try:
        mx = float(mx)
    except:
        mx = 0.0
    pesos = int(mx)
    cents = int(round((mx - pesos)*100))
    ptxt = numero_a_letras(pesos).upper()
    return f"{ptxt} PESOS {cents:02d}/100 M.N."

def pick_col(row: pd.Series, candidates, contains=None):
    for c in candidates:
        if c in row.index: return row.get(c)
        for col in row.index:
            if str(col).lower() == str(c).lower():
                return row.get(col)
    if contains:
        for col in row.index:
            for frag in contains:
                if frag.lower() in str(col).lower():
                    return row.get(col)
    return None

def parse_money(x) -> float:
    if x is None or x == "": return 0.0
    if isinstance(x,(int,float)): return float(x)
    s = str(x).replace("$","").replace(",","").strip()
    try: return float(s)
    except: return 0.0

# ===== NUEVO: Construir dirección para pagaré =====
def build_address(row):
    calle   = pick_col(row, ["Calle","Domicilio","Direccion"], contains=["calle"])
    noext   = pick_col(row, ["NoExt","Num Ext","Exterior","Número exterior", "número exterior"], contains=["ext"])
    noint   = pick_col(row, ["NoInt","Num Int","Interior","Número interior", "número interior"], contains=["int"])
    col     = pick_col(row, ["Colonia"])
    loc     = pick_col(row, ["Localidad","Poblacion"])
    mun     = pick_col(row, ["Municipio","Sucursal"], contains=["mun"])
    edo     = pick_col(row, ["Estado"])
    cp      = pick_col(row, ["CP","C.P","Codigo Postal"], contains=["cp"])

    parts=[]
    if calle: parts.append(str(calle))
    if noext: parts.append(f"#{noext}")
    if noint: parts.append(f"Int {noint}")
    if col: parts.append(f"Col. {col}")
    if loc: parts.append(loc)
    if mun: parts.append(mun)
    if edo: parts.append(f"Edo. {edo}")
    if cp: parts.append(f"C.P. {cp}")

    return ", ".join([p for p in parts if p]) or ""
# ===== DOCX -> PDF (LibreOffice) =====
def _docx_to_pdf_via_libreoffice(in_path: str, out_dir: str) -> str | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    try:
        env = os.environ.copy()
        env.setdefault("HOME", tempfile.gettempdir())
        cmd = [
            soffice, "--headless", "--nologo", "--norestore",
            "--convert-to", "pdf:writer_pdf_Export",
            "--outdir", out_dir, in_path
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
        pdf_path = Path(out_dir) / (Path(in_path).stem + ".pdf")
        return str(pdf_path) if pdf_path.exists() else None
    except Exception:
        return None

def docx_bytes_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    with tempfile.TemporaryDirectory() as td:
        in_path = str(Path(td) / "input.docx")
        Path(in_path).write_bytes(docx_bytes)
        pdf_path = _docx_to_pdf_via_libreoffice(in_path, td)
        if pdf_path and Path(pdf_path).exists():
            return Path(pdf_path).read_bytes()
    return None

# ====== Renderers ======
def letra_abc(idx):
    import string
    i = int(idx) - 1
    if 0 <= i < 26:
        return f"{string.ascii_lowercase[i]})"
    return f"{idx})"

def render_docx(path_tpl: Path, context: dict) -> bytes:
    tpl = DocxTemplate(str(path_tpl))
    try:
        tpl.jinja_env.filters['letra_abc'] = letra_abc
    except Exception:
        pass
    with tempfile.TemporaryDirectory() as td:
        out_path = Path(td) / "out.docx"
        tpl.render(dict(context))
        tpl.save(out_path)
        return out_path.read_bytes()

def render_convenio_con_imagenes(path_tpl: Path, context: dict,
                                 img_pagos_path=None,
                                 img_amort_path=None,
                                 img_control_path=None,
                                 max_width_mm: float = 100.0) -> bytes:
    tpl = DocxTemplate(str(path_tpl))
    try:
        tpl.jinja_env.filters['letra_abc'] = letra_abc
    except Exception:
        pass
    ctx = dict(context)
    # Insert images with a smaller default width (configurable via max_width_mm)
    if img_pagos_path:
        ctx["imagen_tabla_pagos"] = InlineImage(tpl, img_pagos_path, width=Mm(max_width_mm))
    if img_amort_path:
        ctx["imagen_tabla_amort"] = InlineImage(tpl, img_amort_path, width=Mm(max_width_mm))
    if img_control_path:
        ctx["imagen_control_pagos"] = InlineImage(tpl, img_control_path, width=Mm(max_width_mm))
    with tempfile.TemporaryDirectory() as td:
        out_path = Path(td) / "out.docx"
        tpl.render(ctx)
        tpl.save(out_path)
        return out_path.read_bytes()

# ====== Plantillas (Pagaré ÚNICO + Convenio) ======
TPL_DIR = Path(__file__).parent / "plantillas"
PAGARE_UNICO = TPL_DIR / "PAGARE_30_OCT_TEMPLATE.docx"   # <<-- ÚNICA PLANTILLA
TEMPLATE_CONVENIO = TPL_DIR / "CONVENIO_GRUPALejefinal101.docx"  # tu convenio original

# ====== Contextos ======
def row_to_context(row: pd.Series) -> dict:
    """
    CONTEXTO PARA PAGARÉ INDIVIDUAL CON PLANTILLA ÚNICA
    - Mantiene nombres de campos originales para compatibilidad
    - Agrega DireccionCompleta (DOMICILIO DEL CLIENTE) desde Excel
    """
    ctx = {col: row[col] for col in row.index}

    nombre  = pick_col(row, ["Nombre Cliente","Nombre","Cliente"])
    folio   = pick_col(row, ["Clave Solicitud","Folio","Id","ID"])
    # Monto prioriza campo específico de pagaré; si no, usa CUOTA o Monto
    monto_raw = pick_col(
        row,
        ["Monto Pagaré","Monto Pagare","Monto Pagaré "],
        contains=["monto pagar","pagare","pagaré"]
    ) or pick_col(row, ["CUOTA","Monto Autorizado","Monto","Importe","Crédito","Credito"], contains=["cuota","monto","importe","credito"])
    monto_val = parse_money(monto_raw)

    # Dirección completa armada del Excel
    direccion_cliente = build_address(row)

    # Campos para plantilla
    ctx["Nombre"]            = str(nombre or "")
    ctx["Folio"]             = str(folio or "")
    ctx["CUOTA"]             = float(monto_val)
    ctx["CUOTA_FORMAT"]      = f"{monto_val:,.2f}"
    ctx["CUOTA_LETRAS"]      = monto_en_letras(monto_val)
    ctx["DireccionCompleta"] = direccion_cliente         # <<-- NUEVO (para usar en: Domicilio:)
    ctx["DireccionSucursal"] = direccion_cliente         # si tu plantilla usa este nombre, lo mapeamos también
    ctx["FechaHoy"]          = fecha_hoy_es()

    # Compat (si tu plantilla antigua leía Municipio/Sucursal)
    suc_raw = pick_col(row, ["Sucursal","Municipio"]) or ""
    ctx["Municipio"] = str(suc_raw)

    return ctx

def detectar_grupos_kgrupal(df: pd.DataFrame):
    grupos = []
    en_grupo = False
    inicio = None
    for pos, (_, row) in enumerate(df.iterrows()):
        prod = pick_col(row, ["Producto"]) or ""
        es_k = "KGRUPAL" in str(prod).upper()
        if es_k and not en_grupo:
            en_grupo = True
            inicio = pos
        elif (not es_k) and en_grupo:
            grupos.append((inicio, pos - 1))
            en_grupo = False
            inicio = None
    if en_grupo:
        grupos.append((inicio, len(df) - 1))
    return grupos

def crear_contexto_grupal(grupo_df: pd.DataFrame, datos_grupo: dict, montos_antecedentes=None) -> dict:
    """
    Crea contexto del convenio:
    - Suma de pagarés (TotalGrupo)
    - Suma de antecedentes (TotalAntecedentes)
    - Lista de Integrantes con Monto y MontoAntecedente
    - (Se conserva tu lógica original)
    """
    montos_antecedentes = montos_antecedentes or {}
    total_pagare = 0.0
    total_antecedentes = 0.0
    integrantes = []

    for _, row in grupo_df.iterrows():
        nombre = pick_col(row, ["Nombre Cliente","Nombre"]) or ""
        folio  = pick_col(row, ["Clave Solicitud","Folio"]) or ""

        monto_raw = pick_col(row, ["Monto Pagaré","Monto Pagare","Monto Pagaré "], contains=["monto pagar"]) \
                    or pick_col(row, ["CUOTA","Monto","Importe"], contains=["cuota","monto","importe"])
        monto_pagare = parse_money(monto_raw)

        if str(folio) in montos_antecedentes:
            monto_ant = float(montos_antecedentes[str(folio)] or 0)
        else:
            ant_raw = pick_col(row, ["Monto Dispuesto"], contains=["monto dispuesto"])
            monto_ant = parse_money(ant_raw)

        total_pagare += monto_pagare
        total_antecedentes += monto_ant

        integrantes.append({
            "Nombre": str(nombre),
            "Folio": str(folio),
            "Monto": monto_pagare,
            "Monto_FORMAT": f"{monto_pagare:,.2f}",
            "MontoAntecedente": monto_ant,
            "MontoAntecedente_FORMAT": f"{monto_ant:,.2f}",
        })

    lista_integrantes = ", ".join([i["Nombre"] for i in integrantes])

    ctx = {
        "GrupoNombre": datos_grupo.get("nombre_grupo", ""),
        "Integrantes": integrantes,
        "lista_integrantes": lista_integrantes,

        "TotalGrupo": total_pagare,
        "TotalGrupo_FORMAT": f"{total_pagare:,.2f}",
        "TotalGrupo_LETRAS": monto_en_letras(total_pagare),

        "TotalAntecedentes": total_antecedentes,
        "TotalAntecedentes_FORMAT": f"{total_antecedentes:,.2f}",
        "TotalAntecedentes_LETRAS": monto_en_letras(total_antecedentes),

        "FechaHoy": fecha_hoy_es(),
        "FechaFirma": datos_grupo.get("fecha_firma", fecha_hoy_es()),
        "Presidenta": datos_grupo.get("presidenta", ""),
        "Secretaria": datos_grupo.get("secretaria", ""),
        "Tesorera": datos_grupo.get("tesorera", ""),
    }
    return ctx

# ========================= UI =========================
st.set_page_config(page_title="Generador de documentos", page_icon="📄", layout="wide")
st.title("📄 Generador de Pagarés y Convenios Grupales (Plantilla Única Pagaré)")

# Diagnóstico LibreOffice (útil en Cloud)
_soffice = shutil.which("soffice") or shutil.which("libreoffice")
st.sidebar.caption(f"PDF backend: {'LibreOffice encontrado ✅' if _soffice else 'LibreOffice NO disponible ❌'}")

# Subida de Excel (opcional)
excel_file = st.file_uploader("Excel de entrada (.xlsx) (opcional)", type=["xlsx"], accept_multiple_files=False)

# Cargar DF si hay excel
df = pd.DataFrame()
if excel_file:
    with st.spinner("Leyendo Excel..."):
        df = read_excel(excel_file).fillna("").reset_index(drop=True)

# Pestañas SIEMPRE visibles
tab1, tab2 = st.tabs(["📄 PAGARÉS INDIVIDUALES", "👥 CONVENIO GRUPAL"])

# ============ TAB 1: Pagarés individuales ============

with tab1:
    st.subheader("Generar Pagarés Individuales")
    modo = st.radio("Origen de datos", ["Captura manual", "Desde Excel"], horizontal=True)

    # ---- CAPTURA MANUAL ----
    if "manual_pagares" not in st.session_state:
        st.session_state.manual_pagares = []

    if modo == "Captura manual":
        st.info("Captura uno o más pagarés manualmente. Puedes editar la lista antes de generar.")
        gen_pdf_manual = st.checkbox("📄 Generar también PDF de cada pagaré", value=False, key="gen_pdf_manual")

        with st.form("form_manual_pagare", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                nombre_m = st.text_input("Nombre del Cliente *")
                folio_m = st.text_input("Folio / Clave Solicitud *")
            with col2:
                monto_m = st.number_input("Monto Pagaré *", min_value=0.0, step=100.0, format="%.2f")
                fecha_m = st.text_input("Fecha (por defecto hoy)", value=fecha_hoy_es())

            # NUEVO: domicilio completo manual (si no viene de Excel)
            direccion_m = st.text_area("Domicilio del cliente (se imprime como {{DireccionCompleta}}):", value="", placeholder="Calle, #Ext, Int, Colonia, Localidad, Municipio, Estado, C.P. 00000")

            add = st.form_submit_button("➕ Añadir a la lista")
            if add:
                if not nombre_m or not folio_m:
                    st.warning("Completa al menos: Nombre y Folio.")
                else:
                    ctx = {
                        "Nombre": nombre_m,
                        "Folio": folio_m,
                        "CUOTA": float(monto_m),
                        "CUOTA_FORMAT": f"{monto_m:,.2f}",
                        "CUOTA_LETRAS": monto_en_letras(float(monto_m)),
                        "FechaHoy": fecha_m or fecha_hoy_es(),
                        # Tag usado por la plantilla única
                        "DireccionCompleta": (direccion_m or "").strip(),
                        # Compat opcional si la plantilla usa DireccionSucursal
                        "DireccionSucursal": (direccion_m or "").strip(),
                    }
                    st.session_state.manual_pagares.append(ctx)
                    st.success("Añadido ✅")

        # --- Tabla editable + eliminar filas ---
        if st.session_state.manual_pagares:
            ed_df = pd.DataFrame(st.session_state.manual_pagares)

            eliminar_col = "_eliminar"
            if eliminar_col not in ed_df.columns:
                ed_df[eliminar_col] = False

            ed_df = st.data_editor(
                ed_df,
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    eliminar_col: st.column_config.CheckboxColumn(
                        "Eliminar",
                        help="Marca las filas que quieras borrar y pulsa '🗑️ Eliminar seleccionados'",
                        default=False
                    )
                },
                column_order=[eliminar_col] + [c for c in ed_df.columns if c != eliminar_col],
                key="editor_manual_pagares"
            )

            bcol1, bcol2, bcol3 = st.columns([1,1,3])
            with bcol1:
                do_delete = st.button("🗑️ Eliminar seleccionados")
            with bcol2:
                do_clear = st.button("🧹 Vaciar todo")

            if do_clear:
                st.session_state.manual_pagares = []
                st.success("Lista vaciada ✅")
            else:
                if do_delete:
                    restantes = ed_df.loc[~ed_df[eliminar_col]].drop(columns=[eliminar_col], errors="ignore")
                    st.session_state.manual_pagares = restantes.to_dict(orient="records")
                    st.success("Filas seleccionadas eliminadas ✅")
                else:
                    persist = ed_df.drop(columns=[eliminar_col], errors="ignore")
                    st.session_state.manual_pagares = persist.to_dict(orient="records")

            # --- Botón generar (usa PLANTILLA ÚNICA) ---
            if st.button("🚀 Generar Pagarés (Manual)"):
                with st.spinner("Generando pagarés (manual)..."):
                    try:
                        tmp_root = Path(tempfile.mkdtemp(prefix="pagares_manual_"))
                        total, errors = 0, []

                        for i, ctx in enumerate(st.session_state.manual_pagares):
                            tpl_path = PAGARE_UNICO
                            if not tpl_path or not tpl_path.exists():
                                errors.append((i, "No se encontró la plantilla única de pagaré."))
                                continue
                            try:
                                docx_bytes = render_docx(tpl_path, ctx)
                                folder = tmp_root / safe_name(f"{ctx.get('Nombre','SIN_NOMBRE')}")
                                folder.mkdir(parents=True, exist_ok=True)
                                docx_name = safe_name(f"{ctx.get('Folio','')}_{ctx.get('Nombre','')}.docx")
                                pdf_name  = Path(docx_name).with_suffix(".pdf").name

                                (folder / docx_name).write_bytes(docx_bytes)

                                if gen_pdf_manual:
                                    pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes)
                                    if pdf_bytes:
                                        (folder / pdf_name).write_bytes(pdf_bytes)

                                total += 1
                            except Exception as e:
                                errors.append((i, f"Error renderizando: {e}"))

                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for root, _, files in os.walk(tmp_root):
                                for f in files:
                                    full_path = Path(root)/f
                                    rel = full_path.relative_to(tmp_root)
                                    zf.write(full_path, arcname=str(rel))
                        zip_buffer.seek(0)

                        st.success(f"✅ Generados {total} pagaré(s) manual(es)")
                        st.download_button("⬇️ Descargar ZIP", zip_buffer, "pagares_manual.zip", "application/zip")

                        if errors:
                            st.warning("Avisos:")
                            for idx, msg in errors[:200]:
                                st.text(f"{idx}: {msg}")
                    except Exception as e:
                        st.exception(e)

    # ---- DESDE EXCEL ----
    else:
        if df.empty:
            st.warning("Sube un Excel para generar desde archivo.")
            st.button("🚀 Generar Pagarés (Excel)", disabled=True)
        else:
            gen_pdf_excel = st.checkbox("📄 Generar también PDF de cada pagaré", value=False, key="gen_pdf_excel")

            # Opciones avanzadas: ahora solo permiten sobreescritura de domicilio si quieres
            with st.expander("⚙️ Opciones (domicilio)", expanded=False):
                direccion_forzada = st.text_area(
                    "Domicilio (DirecciónCompleta) para usar en TODOS (dejar vacío para usar el del Excel)",
                    value="", key="dir_override"
                )
                direccion_aplicar_todos = st.checkbox(
                    "Aplicar esta dirección a TODOS los pagarés",
                    value=False, key="force_addr_all"
                )

            if st.button("🚀 Generar Pagarés (Excel)"):
                with st.spinner("Generando pagarés desde Excel..."):
                    try:
                        tmp_root = Path(tempfile.mkdtemp(prefix="pagares_excel_"))
                        total, errors = 0, []
                        grupos_kgrupal = detectar_grupos_kgrupal(df)

                        for i, row in df.iterrows():
                            # Excluir filas pertenecientes a grupos KGRUPAL (se generan en la pestaña de convenio)
                            if any(start <= i <= end for (start, end) in grupos_kgrupal):
                                continue

                            ctx = row_to_context(row)

                            # Forzar dirección si así se marcó
                            if direccion_aplicar_todos and direccion_forzada.strip():
                                ctx["DireccionCompleta"] = direccion_forzada.strip()
                                ctx["DireccionSucursal"] = direccion_forzada.strip()  # compat

                            # ÚNICA PLANTILLA
                            tpl_path = PAGARE_UNICO
                            if not tpl_path or not tpl_path.exists():
                                errors.append((ctx.get('Folio','?'), "No se encontró la plantilla única de pagaré."))
                                continue

                            try:
                                docx_bytes = render_docx(tpl_path, ctx)
                                folder = tmp_root / safe_name(f"{ctx.get('Nombre','SIN')}")
                                folder.mkdir(parents=True, exist_ok=True)
                                docx_name = safe_name(f"{ctx.get('Folio','')}_{ctx.get('Nombre','')}.docx")
                                pdf_name  = Path(docx_name).with_suffix(".pdf").name

                                (folder / docx_name).write_bytes(docx_bytes)

                                if gen_pdf_excel:
                                    pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes)
                                    if pdf_bytes:
                                        (folder / pdf_name).write_bytes(pdf_bytes)

                                total += 1
                            except Exception as e:
                                errors.append((ctx.get('Folio','?'), f"Render error: {e}"))

                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for root, _, files in os.walk(tmp_root):
                                for f in files:
                                    full_path = Path(root)/f
                                    rel = full_path.relative_to(tmp_root)
                                    zf.write(full_path, arcname=str(rel))
                        zip_buffer.seek(0)

                        st.success(f"✅ Generados {total} pagaré(s) desde Excel")
                        st.download_button("⬇️ Descargar ZIP", zip_buffer, "pagares_excel.zip", "application/zip")

                        if errors:
                            st.warning("Avisos:")
                            for folio, msg in errors[:200]:
                                st.text(f"{folio}: {msg}")
                    except Exception as e:
                        st.exception(e)
# ============ TAB 2: Convenios Grupales (KGRUPAL) ============

with tab2:
    st.subheader("Generar Convenios Grupales (KGRUPAL)")
    st.info("Esta sección genera los convenios grupales con los pagarés individuales de cada integrante.")

    if df.empty:
        st.warning("Sube un Excel primero para detectar grupos KGRUPAL.")
    else:
        grupos = detectar_grupos_kgrupal(df)
        if not grupos:
            st.info("No se detectaron productos 'KGRUPAL' en tu Excel.")
        else:
            st.success(f"Detectados {len(grupos)} grupo(s) KGRUPAL en tu archivo.")

            # Mostrar selector para elegir un único grupo a procesar
            opciones = [f"Grupo {i+1} (filas {start+1}-{end+1})" for i,(start,end) in enumerate(grupos)]
            sel = st.selectbox("Selecciona un grupo para procesar:", opciones, key="select_grupo_kgrupal")
            gidx = opciones.index(sel)
            ini, fin = grupos[gidx]
            grupo_df = df.iloc[ini:fin+1]
            primera = grupo_df.iloc[0]
            nombre_inicial = pick_col(primera, ["Grupo","Nombre del Grupo","Nombre grupo"]) or f"GRUPO_{gidx+1}"

            # Permitir editar el nombre del grupo antes de generar el convenio
            st.divider()
            nombre_grupo = st.text_input(f"Nombre del Grupo (Grupo {gidx+1})", value=nombre_inicial, key=f"grupo_nombre_{gidx}")
            st.markdown(f"### 🧾 Grupo {gidx+1}: **{nombre_grupo}** ({ini} → {fin})")
            st.dataframe(grupo_df, hide_index=True, use_container_width=True)

            # Campos extra para convenio
            datos_grupo = {
                "nombre_grupo": nombre_grupo,
                "fecha_firma": st.text_input(f"Fecha firma grupo {nombre_grupo}", value=fecha_hoy_es(), key=f"fecha_{gidx}"),
                "presidenta": st.text_input(f"Presidenta grupo {nombre_grupo}", key=f"pres_{gidx}"),
                "secretaria": st.text_input(f"Secretaria grupo {nombre_grupo}", key=f"secr_{gidx}"),
                "tesorera": st.text_input(f"Tesorera grupo {nombre_grupo}", key=f"tes_{gidx}")
            }

            gen_pdf_conv = st.checkbox(f"📄 Generar PDF (grupo {nombre_grupo})", value=False, key=f"pdf_conv_{gidx}")

            # Imágenes anexas para el convenio (se insertarán en la plantilla)
            st.markdown("### 📎 Archivos adicionales (imágenes)")
            tabla_pagos = st.file_uploader("Tabla de pagos concentrada (imagen)", type=["png","jpg","jpeg"], key=f"pagos_{gidx}")
            tabla_amort = st.file_uploader("Tabla de amortización (imagen)", type=["png","jpg","jpeg"], key=f"amort_{gidx}")
            control_pagos = st.file_uploader("Control de pagos (imagen)", type=["png","jpg","jpeg"], key=f"control_{gidx}")

            if st.button(f"🚀 Generar Convenio y Pagarés (Grupo {nombre_grupo})", key=f"btn_generar_grupo_{gidx}"):
                with st.spinner(f"Generando documentos para {nombre_grupo}..."):
                    try:
                        tmp_dir = Path(tempfile.mkdtemp(prefix=f"grupo_{gidx}_"))
                        errors = []

                        # ====== Generar pagarés individuales del grupo ======
                        for _, row in grupo_df.iterrows():
                            ctx = row_to_context(row)

                            tpl_path = PAGARE_UNICO
                            if not tpl_path.exists():
                                errors.append(f"No se encontró la plantilla única de pagaré para {ctx.get('Nombre','')}")
                                continue
                            try:
                                docx_bytes = render_docx(tpl_path, ctx)
                                fname = f"{safe_name(ctx.get('Folio',''))}_{safe_name(ctx.get('Nombre',''))}.docx"
                                (tmp_dir / fname).write_bytes(docx_bytes)

                                if gen_pdf_conv:
                                    pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes)
                                    if pdf_bytes:
                                        (tmp_dir / Path(fname).with_suffix(".pdf")).write_bytes(pdf_bytes)
                            except Exception as e:
                                errors.append(f"Error renderizando pagaré {ctx.get('Folio','')} - {e}")

                        # ====== Generar convenio grupal ======
                        ctx_conv = crear_contexto_grupal(grupo_df, datos_grupo)

                        tpl_conv = TEMPLATE_CONVENIO
                        if not tpl_conv.exists():
                            st.error("❌ No se encontró la plantilla del convenio grupal.")
                            st.stop()

                        # Escribir imágenes temporales y pasarlas al renderer
                        with tempfile.TemporaryDirectory() as td_imgs:
                            p1=p2=p3=None
                            if tabla_pagos is not None:
                                p1 = str(Path(td_imgs)/("pagos"+Path(tabla_pagos.name).suffix)); Path(p1).write_bytes(tabla_pagos.getvalue())
                            if tabla_amort is not None:
                                p2 = str(Path(td_imgs)/("amort"+Path(tabla_amort.name).suffix)); Path(p2).write_bytes(tabla_amort.getvalue())
                            if control_pagos is not None:
                                p3 = str(Path(td_imgs)/("control"+Path(control_pagos.name).suffix)); Path(p3).write_bytes(control_pagos.getvalue())

                            docx_conv_bytes = render_convenio_con_imagenes(tpl_conv, ctx_conv, img_pagos_path=p1, img_amort_path=p2, img_control_path=p3)
                            conv_name = f"CONVENIO_{safe_name(nombre_grupo)}.docx"
                            (tmp_dir / conv_name).write_bytes(docx_conv_bytes)

                            if gen_pdf_conv:
                                pdf_conv_bytes = docx_bytes_to_pdf_bytes(docx_conv_bytes)
                                if pdf_conv_bytes:
                                    (tmp_dir / Path(conv_name).with_suffix(".pdf")).write_bytes(pdf_conv_bytes)

                        if gen_pdf_conv:
                            pdf_conv_bytes = docx_bytes_to_pdf_bytes(docx_conv_bytes)
                            if pdf_conv_bytes:
                                (tmp_dir / Path(conv_name).with_suffix(".pdf")).write_bytes(pdf_conv_bytes)

                        # ====== Empaquetar ZIP ======
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                            for f in tmp_dir.glob("**/*"):
                                if f.is_file():
                                    zf.write(f, arcname=f.name)
                        zip_buffer.seek(0)

                        st.success(f"✅ Convenio y pagarés generados para {nombre_grupo}")
                        st.download_button(
                            "⬇️ Descargar ZIP del grupo",
                            zip_buffer,
                            file_name=f"{safe_name(nombre_grupo)}.zip",
                            mime="application/zip",
                            key=f"zip_{gidx}"
                        )

                        if errors:
                            st.warning("Avisos:")
                            for e in errors:
                                st.text(e)

                    except Exception as e:
                        st.exception(e)

# ============ FIN DEL ARCHIVO ============
st.caption("Versión final 2025-10 - Plantilla única Pagaré + Dirección desde Excel | Kapitaliza 🧾")
