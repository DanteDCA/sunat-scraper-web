import streamlit as st
import pandas as pd
import time
import random
import queue
import threading
from io import BytesIO

import requests
from bs4 import BeautifulSoup

try:
    from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx
except ImportError:
    add_script_run_ctx = None
    get_script_run_ctx = None

if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "hubo_error" not in st.session_state:
    st.session_state.hubo_error = False

BATCH_SIZE  = 50   # sin Chrome podemos subir el lote
PAUSA_LOTES = 15   # segundos entre lotes
NUM_WORKERS = 8    # workers sin límite de RAM por navegador

_URL_BASE = "https://e-consultaruc.sunat.gob.pe"
_URL_FORM = f"{_URL_BASE}/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"
_URL_POST = f"{_URL_BASE}/cl-ti-itmrconsruc/jcrS00Alias"

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Content-Type": "application/x-www-form-urlencoded",
    "Referer": _URL_FORM,
    "Origin": _URL_BASE,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-PE,es;q=0.9",
}


# ── Sesión HTTP ───────────────────────────────────────────────────────────────
def _nueva_sesion() -> requests.Session:
    s = requests.Session()
    s.headers.update(_HEADERS)
    try:
        s.get(_URL_FORM, timeout=12)   # obtiene cookie de sesión
    except Exception:
        pass
    return s


# ── Parser HTML ───────────────────────────────────────────────────────────────
def _parsear_campos(html: str) -> dict:
    """
    Lee la página de resultados SUNAT y devuelve un dict {label: valor}.
    Estructura del HTML: cada div.list-group-item tiene un div.row con pares
    col-label (h4 terminando en ':') / col-valor (p o h4).
    """
    soup = BeautifulSoup(html, "html.parser")
    campos: dict = {}

    for item in soup.find_all("div", class_="list-group-item"):
        row = item.find("div", class_="row")
        if not row:
            continue
        cols = row.find_all("div", recursive=False)

        i = 0
        while i < len(cols) - 1:
            h4 = cols[i].find("h4", class_="list-group-item-heading")
            if not h4:
                i += 1
                continue

            label = h4.get_text(strip=True)
            if not label.endswith(":"):
                i += 2
                continue

            label = label[:-1].strip()
            val_el = cols[i + 1].find(["p", "h4"])
            # separator="\n" convierte <br> en saltos — importante para detectar
            # "Afecto al Nuevo RUS" embebido en Nombre Comercial
            val = val_el.get_text(separator="\n").strip() if val_el else "-"
            campos[label] = val or "-"
            i += 2

    return campos


def _extraer_resultado(campos: dict) -> dict:
    r = {k: "-" for k in [
        "Razon Social", "Tipo Contribuyente", "Tipo de Documento",
        "Nombre Comercial", "Afecto RUS", "Estado",
    ]}

    # Razón Social: el campo "Número de RUC" tiene valor "RUC - RAZON SOCIAL"
    for k, v in campos.items():
        if "mero de RUC" in k or k == "Número de RUC":
            r["Razon Social"] = v.split(" - ", 1)[1] if " - " in v else v
            break

    for k, v in campos.items():
        primera_linea = v.split("\n")[0].strip() or "-"
        if "Tipo Contribuyente" in k:
            r["Tipo Contribuyente"] = primera_linea
        elif "Tipo de Documento" in k:
            r["Tipo de Documento"] = primera_linea
        elif "Nombre Comercial" in k:
            r["Nombre Comercial"] = v.strip() or "-"   # conserva saltos para detectar RUS embebido
        elif "RUS" in k:                                # "Afecto al Nuevo RUS"
            r["Afecto RUS"] = primera_linea
        elif "Estado" in k:
            r["Estado"] = primera_linea

    # Afecto RUS puede venir embebido en el texto de Nombre Comercial
    if r["Afecto RUS"] == "-" and r["Nombre Comercial"] != "-":
        nom = r["Nombre Comercial"]
        for sep in ("Afecto al Nuevo RUS:", "Afecto al Nuevo RUS"):
            if sep in nom:
                partes = nom.split(sep, 1)
                r["Nombre Comercial"] = partes[0].strip() or "-"
                r["Afecto RUS"]       = partes[1].replace(":", "").strip() or "-"
                break

    return r


# ── Consulta individual ───────────────────────────────────────────────────────
def _consultar_ruc(session: requests.Session, ruc: str) -> tuple:
    """
    Retorna (estado, resultado_dict).
    estado: 'ok' | 'invalido' | 'session_error' | 'error'
    """
    payload = {
        "accion":   "consPorRuc",
        "razSoc":   "",
        "nroRuc":   ruc,
        "nrodoc":   "",
        "token":    "x",
        "contexto": "ti-it",
        "modo":     "1",
        "search1":  ruc,
        "rbtnTipo": "1",
        "tipdoc":   "1",
    }
    try:
        resp = session.post(_URL_POST, data=payload, timeout=15)
        resp.raise_for_status()
    except requests.RequestException:
        return "error", {}

    html = resp.text

    # Error de sesión: el servidor rechaza la request
    if "Pagina de Error" in html or (
        "problema" in html.lower()
        and "list-group-item-heading" not in html
    ):
        return "session_error", {}

    # RUC no encontrado: redirigido al formulario sin datos
    if "list-group-item-heading" not in html:
        return "invalido", {}

    campos = _parsear_campos(html)
    return "ok", _extraer_resultado(campos)


# ── Worker persistente ────────────────────────────────────────────────────────
def _worker(cola, df, lock, contador, total, progreso, status_text, errores_globales):
    session = _nueva_sesion()
    procesados_sesion = 0
    REFRESCAR_CADA = 60     # nueva sesión cada N RUCs por seguridad

    while True:
        item = cola.get()

        if item is None:
            cola.task_done()
            break

        index, ruc = item
        resultado = {k: "-" for k in [
            "Razon Social", "Tipo Contribuyente", "Tipo de Documento",
            "Nombre Comercial", "Afecto RUS", "Estado",
        ]}

        try:
            if procesados_sesion > 0 and procesados_sesion % REFRESCAR_CADA == 0:
                session = _nueva_sesion()

            for intento in range(3):
                estado, res = _consultar_ruc(session, ruc)

                if estado == "session_error":
                    session = _nueva_sesion()
                    time.sleep(1)
                    continue
                if estado == "invalido":
                    resultado["Razon Social"] = "RUC INVÁLIDO"
                    break
                if estado == "error":
                    time.sleep(1.5)
                    if intento == 2:
                        resultado["Razon Social"] = "ERROR DE CONEXIÓN"
                    continue
                # ok
                resultado.update(res)
                break

            procesados_sesion += 1

        except Exception as e:
            resultado["Razon Social"] = "ERROR DE CONEXIÓN"
            with lock:
                errores_globales.append(f"RUC {ruc}: {e}")

        finally:
            with lock:
                for campo, valor in resultado.items():
                    df.at[index, campo] = valor
                contador[0] += 1
            try:
                progreso.progress(contador[0] / total)
                status_text.info(f"⏳ Procesados: {contador[0]}/{total} — RUC: `{ruc}`")
            except Exception:
                pass
            cola.task_done()

        time.sleep(random.uniform(0.3, 0.7))


# ── Interfaz Streamlit ────────────────────────────────────────────────────────
st.set_page_config(page_title="Consultor RUC SUNAT", page_icon="🚀")
st.title("🔍 Extractor Masivo SUNAT")

archivo_subido = st.file_uploader("Sube tu Excel con columna 'RUC'", type=["xlsx"])

if archivo_subido:
    df = pd.read_excel(archivo_subido)
    if "RUC" not in df.columns:
        st.error("❌ Falta la columna 'RUC'")
    else:
        total = len(df)
        st.info(
            f"📋 **{total} RUCs** listos para procesar — "
            f"lotes de {BATCH_SIZE} con pausa de {PAUSA_LOTES}s entre cada uno"
        )

        if st.button("🚀 Iniciar Procesamiento"):
            st.session_state.excel_bytes = None
            st.session_state.hubo_error  = False

            df["RUC"] = (
                df["RUC"].astype(str)
                .str.replace(r"[^\d]", "", regex=True)
                .str.zfill(11)
            )

            for col in ["Razon Social", "Tipo Contribuyente", "Tipo de Documento",
                        "Nombre Comercial", "Afecto RUS", "Estado"]:
                if col not in df.columns:
                    df[col] = "-"

            progreso    = st.progress(0)
            status_text = st.empty()
            status_text.info("🔧 Iniciando workers...")

            try:
                cola             = queue.Queue()
                lock             = threading.Lock()
                contador         = [0]
                errores_globales = []
                ctx              = get_script_run_ctx() if get_script_run_ctx else None

                hilos = []
                for _ in range(NUM_WORKERS):
                    t = threading.Thread(
                        target=_worker,
                        args=(cola, df, lock, contador, total, progreso,
                              status_text, errores_globales),
                        daemon=True,
                    )
                    if add_script_run_ctx and ctx:
                        add_script_run_ctx(t, ctx)
                    hilos.append(t)
                    t.start()

                todos_rucs = [(i, row["RUC"].strip()) for i, row in df.iterrows()]
                lotes = [todos_rucs[i:i + BATCH_SIZE]
                         for i in range(0, len(todos_rucs), BATCH_SIZE)]

                for num_lote, lote in enumerate(lotes, start=1):
                    for item in lote:
                        cola.put(item)
                    cola.join()

                    if num_lote < len(lotes):
                        for seg in range(PAUSA_LOTES, 0, -1):
                            try:
                                status_text.warning(
                                    f"⏸️ Lote {num_lote}/{len(lotes)} listo — "
                                    f"pausa anti-bloqueo: {seg}s..."
                                )
                            except Exception:
                                pass
                            time.sleep(1)

                for _ in range(NUM_WORKERS):
                    cola.put(None)
                for t in hilos:
                    t.join()

                if errores_globales:
                    st.session_state.hubo_error = True
                    st.warning(f"⚠️ {len(errores_globales)} error(es) durante el proceso.")

                status_text.success("✅ ¡Procesamiento terminado!")

            except Exception as e:
                st.error(f"❌ Error fatal: {e}")
                st.session_state.hubo_error = True

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False)
            st.session_state.excel_bytes = output.getvalue()

        if st.session_state.excel_bytes:
            label = (
                "📥 Descargar Resultados Parciales"
                if st.session_state.hubo_error
                else "📥 Descargar Resultados Finales"
            )
            st.download_button(
                label=label,
                data=st.session_state.excel_bytes,
                file_name="sunat_resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
