"""
Control de expedientes – Streamlit + Polars + Google Sheets
Ejecutar:  streamlit run app.py
"""
from __future__ import annotations

import json
from datetime import datetime

import gspread
import polars as pl
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

st.set_page_config(page_title="Control de expedientes", page_icon="📋", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
BASE_COLS = ["Expediente", "Banco", "Registro", "Observaciones"]
MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
         "AGOSTO", "SETIEMBRE", "SEPTIEMBRE", "SEP", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


# ───────────────────────────── Conexión a Google ─────────────────────────────
@st.cache_resource(show_spinner=False)
def get_client(creds_json: str) -> gspread.Client:
    info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(ttl=300, show_spinner=False)
def open_sheet(creds_json: str, url: str) -> gspread.Spreadsheet:
    return get_client(creds_json).open_by_url(url)


@st.cache_data(ttl=45, show_spinner=False)
def load_values(creds_json: str, url: str, ws_title: str) -> list[list[str]]:
    return open_sheet(creds_json, url).worksheet(ws_title).get_all_values()


def get_credentials_json() -> str | None:
    """Prioridad: archivo subido > texto pegado > st.secrets."""
    up = st.session_state.get("cred_file")
    if up is not None:
        return up.getvalue().decode("utf-8")
    pasted = st.session_state.get("cred_text", "").strip()
    if pasted:
        return pasted
    try:
        if "gcp_service_account" in st.secrets:
            return json.dumps(dict(st.secrets["gcp_service_account"]))
    except Exception:
        pass
    return None


# ───────────────────────────── Polars: datos ─────────────────────────────
def build_df(values: list[list[str]], header_mode: str) -> tuple[pl.DataFrame, bool]:
    """Convierte los valores crudos de la hoja en un DataFrame de Polars.
    Se conserva `_fila` (número de fila real en la hoja) para poder editar."""
    if not values:
        return pl.DataFrame(schema={"_fila": pl.Int64, **{c: pl.Utf8 for c in BASE_COLS}}), False

    first = (values[0][0] if values[0] else "").strip().lower()
    if header_mode == "Sí":
        has_header = True
    elif header_mode == "No":
        has_header = False
    else:  # Auto
        has_header = first.startswith("expediente")

    ncols = max(max(len(r) for r in values), len(BASE_COLS))
    names = list(BASE_COLS)
    for i in range(len(BASE_COLS), ncols):
        title = ""
        if has_header and i < len(values[0]):
            title = values[0][i].strip()
        title = title or f"Col {rowcol_to_a1(1, i + 1)[:-1]}"
        while title in names:
            title += "_"
        names.append(title)

    filas: list[int] = []
    cols: dict[str, list[str]] = {n: [] for n in names}
    for idx, row in enumerate(values):
        if has_header and idx == 0:
            continue
        row = [str(c).strip() for c in row] + [""] * (ncols - len(row))
        if not any(row):
            continue
        filas.append(idx + 1)
        for n, v in zip(names, row):
            cols[n].append(v)

    schema = {"_fila": pl.Int64, **{n: pl.Utf8 for n in names}}
    return pl.DataFrame({"_fila": filas, **cols}, schema=schema), has_header


def enrich(df: pl.DataFrame, label_ok: str) -> pl.DataFrame:
    """Añade banderas booleanas usadas por las métricas."""
    ok = label_ok.strip().upper()
    return df.with_columns(
        (pl.col("Expediente") != "").alias("_tiene_exp"),
        (pl.col("Registro").str.to_uppercase() == ok).alias("_subido"),
    ).with_columns(
        (pl.col("_tiene_exp") & pl.col("_subido")).alias("_ok"),
        (pl.col("_tiene_exp") & ~pl.col("_subido")).alias("_pend"),
        (~pl.col("_tiene_exp") & (pl.col("Observaciones") != "")).alias("_sin_exp"),
        (pl.col("_tiene_exp") & pl.col("Expediente").is_duplicated()).alias("_dup"),
    )


def metrics(df: pl.DataFrame) -> dict:
    total = int(df["_tiene_exp"].sum())
    ok = int(df["_ok"].sum())
    return {
        "total": total,
        "ok": ok,
        "pend": int(df["_pend"].sum()),
        "sin_exp": int(df["_sin_exp"].sum()),
        "dup": int(df.filter(pl.col("_tiene_exp"))["_dup"].sum()),
        "obs": int((df["Observaciones"] != "").sum()),
        "pct": (ok / total * 100) if total else 0.0,
    }


def by_bank(df: pl.DataFrame) -> pl.DataFrame:
    return (
        df.filter(pl.col("_tiene_exp"))
        .with_columns(pl.col("Banco").str.to_uppercase().replace("", "SIN BANCO").alias("Banco"))
        .group_by("Banco")
        .agg(
            pl.len().alias("Total"),
            pl.col("_ok").sum().alias("Subidos"),
            pl.col("_pend").sum().alias("Pendientes"),
        )
        .with_columns((pl.col("Subidos") / pl.col("Total") * 100).round(1).alias("% Avance"))
        .sort("Total", descending=True)
    )


# ───────────────────────────── Escritura ─────────────────────────────
def write_cells(ws: gspread.Worksheet, updates: list[tuple[int, int, str]]) -> None:
    """updates = [(fila, columna, valor)]"""
    if not updates:
        return
    ws.batch_update(
        [{"range": rowcol_to_a1(r, c), "values": [[v]]} for r, c, v in updates],
        value_input_option="USER_ENTERED",
    )
    load_values.clear()

def hex_to_rgb(hex_color: str) -> dict:
    h = str(hex_color).lstrip('#')
    if len(h) == 6:
        try:
            return {"red": int(h[0:2], 16) / 255.0, "green": int(h[2:4], 16) / 255.0, "blue": int(h[4:6], 16) / 255.0}
        except ValueError:
            pass
    return {"red": 1.0, "green": 1.0, "blue": 1.0}


def apply_colors_to_sheet(ws: gspread.Worksheet, df: pl.DataFrame, color_dict: dict) -> None:
    requests = []
    ncols = len([col for col in df.columns if col != "_fila"])
    for row in df.iter_rows(named=True):
        reg = str(row.get("Registro", "")).strip().upper()
        c = color_dict.get(reg, "")
        fila = int(row["_fila"])
        rgb = hex_to_rgb(c) if c and c.startswith("#") else {"red": 1.0, "green": 1.0, "blue": 1.0}
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": fila - 1,
                    "endRowIndex": fila,
                    "startColumnIndex": 0,
                    "endColumnIndex": ncols,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": rgb}},
                "fields": "userEnteredFormat.backgroundColor",
            }
        })
    if requests:
        ws.spreadsheet.batch_update({"requests": requests})


CONFIG_SHEET_NAME = "Configuraciones"


def is_config_sheet(title: str) -> bool:
    return title.strip().upper() in ["CONFIGURACIONES", "CONFIGURACION", "CONFIG"]


def get_default_config() -> list[dict[str, str]]:
    return [
        {"Valor": "PRESENTADO", "Color Hex": "#ead1dc", "Completado": "No"},
        {"Valor": "SUBIDO", "Color Hex": "#c9daf8", "Completado": "Sí"},
    ]


def load_config_from_sheet(sh: gspread.Spreadsheet) -> tuple[list[dict[str, str]], str]:
    default_cfg = get_default_config()
    default_ok = "SUBIDO"
    try:
        ws_list = [w for w in sh.worksheets() if is_config_sheet(w.title)]
        if not ws_list:
            return default_cfg, default_ok
        ws = ws_list[0]
        vals = ws.get_all_values()
        if not vals or len(vals) <= 1:
            return default_cfg, default_ok

        start_row = 1 if any("estado" in c.lower() or "valor" in c.lower() for c in vals[0]) else 0
        cfg_rows = []
        found_ok = None
        for r in vals[start_row:]:
            if not r or not r[0].strip():
                continue
            estado = r[0].strip().upper()
            color = r[1].strip() if len(r) > 1 and r[1].strip() else "#c9daf8"
            if not color.startswith("#") and len(color) == 6:
                color = "#" + color
            comp_str = r[2].strip().lower() if len(r) > 2 and r[2].strip() else ""
            is_comp = comp_str in ["sí", "si", "true", "1", "completado", "yes"] or (estado == "SUBIDO" and not comp_str)
            if is_comp and not found_ok:
                found_ok = estado
            cfg_rows.append({"Valor": estado, "Color Hex": color, "Completado": "Sí" if is_comp else "No"})

        if cfg_rows:
            return cfg_rows, (found_ok or "SUBIDO")
    except Exception:
        pass
    return default_cfg, default_ok


def save_config_to_sheet(sh: gspread.Spreadsheet, cfg_list: list[dict[str, str]], label_ok: str) -> None:
    ws = None
    for w in sh.worksheets():
        if is_config_sheet(w.title):
            ws = w
            break
    if ws is None:
        ws = sh.add_worksheet(title=CONFIG_SHEET_NAME, rows="100", cols="10")

    rows = [["Estado", "Color Hex", "Completado", "Muestra Visual"]]
    ok_upper = label_ok.strip().upper()
    for item in cfg_list:
        v = str(item.get("Valor", "")).strip().upper()
        if not v:
            continue
        c = str(item.get("Color Hex", "")).strip()
        comp = "Sí" if v == ok_upper or item.get("Completado", "") in ["Sí", "si", True] else "No"
        rows.append([v, c, comp, v])

    ws.clear()
    ws.update(range_name="A1", values=rows)

    reqs = [
        {
            "repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 4},
                "cell": {"userEnteredFormat": {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.93, "green": 0.93, "blue": 0.93}}},
                "fields": "userEnteredFormat(textFormat,backgroundColor)",
            }
        }
    ]
    for idx, item in enumerate(cfg_list, start=1):
        color_rgb = hex_to_rgb(item.get("Color Hex", "#ffffff"))
        reqs.append({
            "repeatCell": {
                "range": {"sheetId": ws.id, "startRowIndex": idx, "endRowIndex": idx + 1, "startColumnIndex": 3, "endColumnIndex": 4},
                "cell": {"userEnteredFormat": {"backgroundColor": color_rgb, "textFormat": {"bold": True}}},
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        })
    try:
        ws.spreadsheet.batch_update({"requests": reqs})
    except Exception:
        pass
    load_values.clear()


# ───────────────────────────── Sidebar ─────────────────────────────
with st.sidebar:
    st.title("📋 LuDash")
    st.caption("Control de expedientes")

    st.subheader("1. Hoja de Google")
    sheet_url = st.text_input(
        "Link de la hoja de cálculo",
        key="sheet_url",
        value="https://docs.google.com/spreadsheets/d/1rYVukt9umQEuXnaIesX5OqXCNOHNgz1ZNXLxUqDEPSc/edit?usp=sharing",
        placeholder="https://docs.google.com/spreadsheets/d/…",
        help="Pega el link completo. La hoja debe estar compartida con el usuario virtual.",
    ).strip()

    with st.expander("🔑 Credenciales de servicio", expanded=False):
        st.file_uploader("Archivo JSON de la key", type="json", key="cred_file")
        st.text_area("…o pega el contenido del JSON", key="cred_text", height=70)
        st.caption("También se leen desde `st.secrets['gcp_service_account']`.")

    creds_json = get_credentials_json()
    ws_title = None
    sh = None
    error = None

    if sheet_url and creds_json:
        try:
            sh = open_sheet(creds_json, sheet_url)
            all_ws = [w.title for w in sh.worksheets()]
            # Filtrar hojas que sean de configuración para los meses
            data_titles = [t for t in all_ws if not is_config_sheet(t)]
            if not data_titles:
                data_titles = all_ws

            default = len(data_titles) - 1
            for i, t in enumerate(data_titles):
                if t.strip().upper().startswith(MESES[datetime.now().month - 1][:3]):
                    default = i

            st.subheader("2. Pestaña del mes")
            ws_title = st.selectbox("Mes activo", data_titles, index=default)

            with st.expander("➕ Crear nueva hoja (mes)"):
                new_month = st.text_input("Nombre de la nueva hoja", placeholder="Ej: Noviembre")
                if st.button("Crear hoja"):
                    if not new_month.strip():
                        st.error("Ingresa un nombre para la hoja.")
                    elif new_month.strip() in all_ws:
                        st.error("La hoja ya existe.")
                    else:
                        new_ws = sh.add_worksheet(title=new_month.strip(), rows="1000", cols="20")
                        new_ws.append_row(BASE_COLS)
                        st.success(f"Hoja '{new_month.strip()}' creada.")
                        load_values.clear()
                        st.rerun()

            st.success(f"Conectado: {sh.title}", icon="✅")
            try:
                st.caption(f"Cuenta: `{json.loads(creds_json)['client_email']}`")
            except Exception:
                pass
        except json.JSONDecodeError:
            error = "El JSON de credenciales no es válido."
        except gspread.SpreadsheetNotFound:
            error = "No se encontró la hoja. ¿La compartiste con el usuario virtual?"
        except Exception as e:
            error = f"No se pudo conectar: {e}"
        if error:
            st.error(error)
    else:
        st.info("Ingresa el link y las credenciales para conectar.")

    # Cargar y gestionar estados en session_state
    if sh is not None and "app_config_list" not in st.session_state:
        cfg_list, loaded_ok = load_config_from_sheet(sh)
        st.session_state.app_config_list = cfg_list
        st.session_state.label_ok = loaded_ok

    if "app_config_list" not in st.session_state:
        st.session_state.app_config_list = get_default_config()
        st.session_state.label_ok = "SUBIDO"

    # Diccionario rápido de colores
    color_dict = {
        str(item.get("Valor", "")).strip().upper(): str(item.get("Color Hex", "")).strip()
        for item in st.session_state.app_config_list
        if str(item.get("Valor", "")).strip()
    }
    label_ok = st.session_state.get("label_ok", "SUBIDO")

    st.divider()
    st.subheader("3. Estados de registro")
    badge_html = " ".join([
        f'<span style="background-color: {item.get("Color Hex", "#eee")}; color: #000; padding: 2px 7px; border-radius: 4px; font-size: 11px; font-weight: bold; border: 1px solid rgba(0,0,0,0.1); margin-right: 4px;">{item.get("Valor")}{" ✓" if item.get("Valor") == label_ok else ""}</span>'
        for item in st.session_state.app_config_list
    ])
    st.markdown(badge_html, unsafe_allow_html=True)
    st.caption("Gestiona o agrega estados en la pestaña **⚙️ Configuración**.")

    if st.button("🔄 Recargar datos", use_container_width=True):
        load_values.clear()
        if sh is not None:
            cfg_list, loaded_ok = load_config_from_sheet(sh)
            st.session_state.app_config_list = cfg_list
            st.session_state.label_ok = loaded_ok
        st.rerun()


# Parámetros generales
header_mode = st.session_state.get("header_mode_select", "Auto")
meta = st.session_state.get("meta_input", 0)

# ───────────────────────────── Contenido ─────────────────────────────
st.title("📋 Control de expedientes")
tab_met, tab_edit, tab_comp, tab_conf, tab_help = st.tabs(
    ["📊 Métricas", "✏️ Editar hoja", "📅 Comparativo mensual", "⚙️ Configuración", "🔑 Instrucciones"]
)

df = None
if sh is not None and ws_title:
    raw = load_values(creds_json, sheet_url, ws_title)
    df_raw, has_header = build_df(raw, header_mode)
    df = enrich(df_raw, label_ok)

# ---------- Métricas ----------
with tab_met:
    if df is None:
        st.info("Conecta una hoja desde el sidebar para ver las métricas.")
    else:
        m = metrics(df)
        st.caption(f"Pestaña **{ws_title}** · encabezado detectado: {'sí' if has_header else 'no'}")
        c = st.columns(5)
        c[0].metric("Expedientes", m["total"])
        c[1].metric(label_ok.title(), m["ok"])
        c[2].metric("Pendientes", m["pend"])
        c[3].metric("% de avance", f"{m['pct']:.1f}%")
        c[4].metric("Con observación", m["obs"])
        st.progress(min(m["pct"] / 100, 1.0), text=f"Avance de registro: {m['pct']:.1f}%")

        if meta:
            p = min(m["total"] / meta, 1.0)
            st.progress(p, text=f"Meta mensual: {m['total']} / {meta} expedientes ({m['total']/meta*100:.0f}%)")

        alerts = []
        if m["sin_exp"]:
            alerts.append(f"⚠️ {m['sin_exp']} fila(s) con observaciones pero **sin número de expediente**.")
        if m["dup"]:
            alerts.append(f"⚠️ {m['dup']} fila(s) con expediente **duplicado**.")
        for a in alerts:
            st.warning(a)

        bank = by_bank(df)
        
        st.subheader("Distribución general por Opciones (Registro)")
        import altair as alt
        
        # Generar datos agrupados por Registro
        reg_df = df.filter(pl.col("_tiene_exp")).with_columns(
            pl.col("Registro").str.to_uppercase().replace("", "VACÍO").alias("Registro")
        ).group_by("Registro").agg(pl.len().alias("Cantidad"))
        reg_pd = reg_df.to_pandas()
        
        # Mapeo de colores dinámico
        domain = []
        range_ = []
        for r_val in reg_pd["Registro"].unique():
            domain.append(r_val)
            range_.append(color_dict.get(r_val, "#1f77b4")) # Azul por defecto si no hay color
        
        pie = alt.Chart(reg_pd).mark_arc(innerRadius=40).encode(
            theta=alt.Theta(field="Cantidad", type="quantitative"),
            color=alt.Color(field="Registro", type="nominal", scale=alt.Scale(domain=domain, range=range_)),
            tooltip=["Registro", "Cantidad"]
        ).properties(height=300)
        
        bar_opts = alt.Chart(reg_pd).mark_bar().encode(
            x=alt.X('Registro:N', sort='-y', title=""),
            y=alt.Y('Cantidad:Q', title="Expedientes"),
            color=alt.Color('Registro:N', scale=alt.Scale(domain=domain, range=range_), legend=None),
            tooltip=["Registro", "Cantidad"]
        ).properties(height=300)

        c_pie, c_bar_opt = st.columns(2)
        with c_pie:
            st.altair_chart(pie, use_container_width=True)
        with c_bar_opt:
            st.altair_chart(bar_opts, use_container_width=True)

        st.divider()

        left, right = st.columns([3, 2])
        with left:
            st.subheader("Avance por banco")
            if bank.height:
                c_subido = color_dict.get(label_ok.strip().upper(), "#3366cc")
                c_pend = color_dict.get("PENDIENTE", "#dc3912")
                try:
                    st.bar_chart(bank, x="Banco", y=["Subidos", "Pendientes"], stack=True, color=[c_subido, c_pend])
                except Exception:
                    st.bar_chart(bank, x="Banco", y=["Subidos", "Pendientes"], stack=True)
        with right:
            st.subheader("Detalle")
            st.dataframe(bank, hide_index=True)

        st.subheader("Pendientes de registro")
        pend = df.filter(pl.col("_pend")).select("_fila", "Expediente", "Banco", "Registro", "Observaciones")
        st.dataframe(pend, hide_index=True)

# ---------- Editar ----------
with tab_edit:
    if df is None:
        st.info("Conecta una hoja desde el sidebar para editarla.")
    else:
        ws = sh.worksheet(ws_title)
        ncols_sheet = len([c for c in df_raw.columns if c != "_fila"])

        st.subheader("Editar registros existentes")
        vista = st.radio(
            "Ver", ["Todos", "Pendientes", "Observación sin expediente", "Duplicados"],
            horizontal=True,
        )
        filtro = {
            "Todos": pl.lit(True),
            "Pendientes": pl.col("_pend"),
            "Observación sin expediente": pl.col("_sin_exp"),
            "Duplicados": pl.col("_dup") & pl.col("_tiene_exp"),
        }[vista]
        buscar = st.text_input("Buscar (expediente, banco u observación)", "")
        view = df.filter(filtro).select(df_raw.columns)
        if buscar:
            q = buscar.lower()
            view = view.filter(
                pl.any_horizontal(pl.col(c).str.to_lowercase().str.contains(q, literal=True)
                                  for c in view.columns if c != "_fila")
            )

        def row_style(row):
            reg = str(row.get("Registro", "")).strip().upper()
            color = color_dict.get(reg, "")
            return [f"background-color: {color}; color: #000000" if color else ""] * len(row)

        view_pd = view.to_pandas()
        styled_view = view_pd.style.apply(row_style, axis=1)

        edited = st.data_editor(
            styled_view,
            hide_index=True,
            disabled=["_fila"],
            column_config={"_fila": st.column_config.NumberColumn("Fila", format="%d")},
            key=f"editor_{ws_title}_{vista}_{buscar}",
        )
        if not isinstance(edited, pl.DataFrame):
            edited = pl.from_pandas(edited)

        # Detectar celdas cambiadas
        cols = [c for c in view.columns if c != "_fila"]
        updates: list[tuple[int, int, str]] = []
        for old, new in zip(view.iter_rows(named=True), edited.iter_rows(named=True)):
            for col in cols:
                if (old[col] or "") != (new[col] or ""):
                    updates.append((old["_fila"], df_raw.columns.index(col), new[col] or ""))
        # df_raw.columns incluye _fila en la posición 0 → índice = posición de columna (1-based)

        cc = st.columns([2, 2, 2])
        if cc[0].button(f"💾 Guardar {len(updates)} cambio(s)", type="primary", disabled=not updates):
            write_cells(ws, updates)
            st.success("Cambios guardados en Google Sheets.")
            st.rerun()
            
        if cc[1].button("🎨 Sincronizar colores al Excel"):
            with st.spinner("Pintando filas en Google Sheets..."):
                # Se envía el dataframe completo actual con la config
                apply_colors_to_sheet(ws, df, color_dict)
            st.success("¡Colores sincronizados!")
            
        cc[2].caption("Presiona Sincronizar colores para aplicarlos al archivo de Google Sheets.")
        st.divider()
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Actualizar pendientes")
            pend_list = df.filter(pl.col("_pend"))
            opciones = {f"{r['Expediente']} · {r['Banco']}": r["_fila"] for r in pend_list.iter_rows(named=True)}
            sel = st.multiselect("Expedientes pendientes", list(opciones))
            
            current_states = [item["Valor"] for item in st.session_state.app_config_list if item.get("Valor")]
            default_state_idx = current_states.index(label_ok) if label_ok in current_states else 0
            sel_estado = st.selectbox("Nuevo Registro", current_states + ["(otro)"], index=default_state_idx)
            n_estado = st.text_input("Especificar otro registro") if sel_estado == "(otro)" else sel_estado
            n_obs = st.text_input("Agregar/Cambiar Observación (Opcional)")

            if st.button(f"✅ Actualizar {len(sel)} expediente(s)", disabled=not sel):
                upd = []
                for s in sel:
                    fila = opciones[s]
                    if n_estado.strip():
                        upd.append((fila, 3, n_estado.strip().upper()))
                    if n_obs.strip():
                        upd.append((fila, 4, n_obs.strip()))
                if upd:
                    write_cells(ws, upd)
                    st.success(f"{len(sel)} expediente(s) actualizados.")
                    st.rerun()

        with col2:
            st.subheader("Agregar nuevo expediente")
            with st.form("nuevo", clear_on_submit=True):
                n_exp = st.text_input("Expediente", placeholder="0000-2026")
                bancos = sorted({b for b in df["Banco"].to_list() if b})
                n_banco = st.selectbox("Banco", bancos + ["(otro)"])
                n_otro = st.text_input("Otro banco (si elegiste '(otro)')")
                sel_n_reg = st.selectbox("Registro", current_states + ["(otro)"], index=default_state_idx)
                n_reg = st.text_input("Especificar otro registro") if sel_n_reg == "(otro)" else sel_n_reg
                n_obs = st.text_input("Observaciones")
                if st.form_submit_button("➕ Agregar al final de la hoja"):
                    if not n_exp.strip():
                        st.error("El expediente es obligatorio.")
                    else:
                        banco = n_otro.strip().upper() if n_banco == "(otro)" else n_banco
                        if n_exp.strip() in df["Expediente"].to_list():
                            st.warning("Ese expediente ya existe; se agregó igualmente.")
                        row = [n_exp.strip(), banco, n_reg.strip(), n_obs.strip()]
                        ws.append_row(row, value_input_option="USER_ENTERED")
                        load_values.clear()
                        st.success("Registro agregado.")
                        st.rerun()

# ---------- Comparativo ----------
with tab_comp:
    if sh is None:
        st.info("Conecta una hoja desde el sidebar.")
    else:
        st.write("Visión Global: Resume los datos de las pestañas (meses) seleccionados.")
        
        titles = [w.title for w in sh.worksheets()]
        data_titles = [t for t in titles if not is_config_sheet(t)]
        if not data_titles:
            data_titles = titles
        selected_months = st.multiselect("Filtrar meses (deja vacío para analizarlos TODOS)", data_titles, default=[])
        
        if st.button("📅 Generar Reporte Global", type="primary"):
            meses_a_leer = selected_months if selected_months else data_titles
            
            all_dfs = []
            filas_mensuales = []
            
            with st.spinner("Procesando datos globales..."):
                for t in meses_a_leer:
                    try:
                        raw = load_values(creds_json, sheet_url, t)
                        if raw:
                            d, _ = build_df(raw, header_mode)
                            if d.height > 0:
                                d_en = enrich(d, label_ok)
                                d_en = d_en.with_columns(pl.lit(t).alias("Mes"))
                                all_dfs.append(d_en)
                                
                                mm = metrics(d_en)
                                filas_mensuales.append({"Mes": t, "Expedientes": mm["total"], "Completados": mm["ok"],
                                                        "Pendientes": mm["pend"], "% Avance": round(mm["pct"], 1)})
                    except Exception:
                        pass
                        
            if not all_dfs:
                st.warning("No se encontraron datos en los meses seleccionados.")
            else:
                global_df = pl.concat(all_dfs, how="vertical_relaxed")
                gm = metrics(global_df)
                
                st.subheader("📊 Resumen Ejecutivo (Global)")
                c = st.columns(4)
                c[0].metric("Total Expedientes", gm["total"])
                c[1].metric(label_ok.title(), gm["ok"])
                c[2].metric("Total Pendientes", gm["pend"])
                c[3].metric("% Avance Global", f"{gm['pct']:.1f}%")
                st.progress(min(gm["pct"] / 100, 1.0))
                
                st.divider()
                
                st.subheader("Evolución Mensual")
                comp = pl.DataFrame(filas_mensuales)
                c_subido = color_dict.get(label_ok.strip().upper(), "#3366cc")
                c_pend = color_dict.get("PENDIENTE", "#dc3912")
                
                c_table, c_chart = st.columns([2, 3])
                with c_table:
                    st.dataframe(comp, hide_index=True)
                with c_chart:
                    try:
                        st.bar_chart(comp, x="Mes", y=["Completados", "Pendientes"], stack=True, color=[c_subido, c_pend])
                    except Exception:
                        st.bar_chart(comp, x="Mes", y=["Completados", "Pendientes"], stack=True)
                
                st.divider()
                
                st.subheader("Distribución Global por Estado (Registro)")
                import altair as alt
                
                reg_df = global_df.filter(pl.col("_tiene_exp")).with_columns(
                    pl.col("Registro").str.to_uppercase().replace("", "VACÍO").alias("Registro")
                ).group_by("Registro").agg(pl.len().alias("Cantidad"))
                reg_pd = reg_df.to_pandas()
                
                domain = []
                range_ = []
                for r_val in reg_pd["Registro"].unique():
                    domain.append(r_val)
                    range_.append(color_dict.get(r_val, "#1f77b4"))
                
                pie = alt.Chart(reg_pd).mark_arc(innerRadius=40).encode(
                    theta=alt.Theta(field="Cantidad", type="quantitative"),
                    color=alt.Color(field="Registro", type="nominal", scale=alt.Scale(domain=domain, range=range_)),
                    tooltip=["Registro", "Cantidad"]
                ).properties(height=350)
                
                bar_opts = alt.Chart(reg_pd).mark_bar().encode(
                    x=alt.X('Registro:N', sort='-y', title=""),
                    y=alt.Y('Cantidad:Q', title="Expedientes"),
                    color=alt.Color('Registro:N', scale=alt.Scale(domain=domain, range=range_), legend=None),
                    tooltip=["Registro", "Cantidad"]
                ).properties(height=350)
                
                c_pie, c_bar_opt = st.columns(2)
                with c_pie:
                    st.altair_chart(pie, use_container_width=True)
                with c_bar_opt:
                    st.altair_chart(bar_opts, use_container_width=True)

# ---------- Configuración ----------
with tab_conf:
    st.subheader("⚙️ Configuración de Estados y Colores")
    st.markdown(
        "Personaliza los estados para clasificar tus expedientes (por ejemplo **PRESENTADO** y **SUBIDO**), "
        "asigna colores visualmente y sincronízalos directamente con la pestaña **Configuraciones** en Google Sheets."
    )

    if sh is None:
        st.warning("⚠️ Conecta tu hoja de Google en el panel lateral para poder sincronizar las configuraciones.")
    else:
        # Botones de acción principales
        col_s1, col_s2, col_s3 = st.columns([2, 2, 2])
        with col_s1:
            if st.button("📤 Guardar en Google Sheets", type="primary", use_container_width=True, help="Sube los estados y colores a la pestaña 'Configuraciones'"):
                with st.spinner("Guardando en la hoja Configuraciones..."):
                    save_config_to_sheet(sh, st.session_state.app_config_list, st.session_state.label_ok)
                st.success("¡Configuración guardada y formateada en la pestaña Configuraciones de Google Sheets!")
        with col_s2:
            if st.button("📥 Cargar desde Google Sheets", use_container_width=True, help="Lee los estados guardados en la pestaña 'Configuraciones'"):
                with st.spinner("Leyendo hoja Configuraciones..."):
                    cfg_list, loaded_ok = load_config_from_sheet(sh)
                    st.session_state.app_config_list = cfg_list
                    st.session_state.label_ok = loaded_ok
                st.success("¡Configuración cargada desde Google Sheets!")
                st.rerun()
        with col_s3:
            if ws_title and df is not None:
                if st.button(f"🎨 Pintar mes actual ({ws_title})", use_container_width=True, help="Aplica los colores configurados a todas las filas en Google Sheets"):
                    with st.spinner("Pintando filas en Google Sheets..."):
                        ws_curr = sh.worksheet(ws_title)
                        apply_colors_to_sheet(ws_curr, df, color_dict)
                    st.success(f"¡Filas del mes {ws_title} pintadas exitosamente!")

    st.divider()

    # Opciones para cuál estado cuenta como completado
    st.subheader("Estados registrados")
    st.caption("Selecciona con el selector de color el tono para cada estado o indica cuál cuenta como completado.")

    current_states = [item["Valor"] for item in st.session_state.app_config_list if item.get("Valor")]
    default_ok_idx = current_states.index(st.session_state.label_ok) if st.session_state.label_ok in current_states else 0

    selected_ok = st.selectbox(
        "🏆 ¿Cuál estado cuenta como COMPLETADO / SUBIDO? (Para métricas y avances)",
        options=current_states,
        index=default_ok_idx,
        help="El estado seleccionado sumará para el porcentaje de avance y expedientes listos."
    )
    if selected_ok != st.session_state.label_ok:
        st.session_state.label_ok = selected_ok
        for item in st.session_state.app_config_list:
            item["Completado"] = "Sí" if item["Valor"] == selected_ok else "No"
        st.rerun()

    # Lista interactiva de tarjetas por cada estado
    eliminar_idx = None
    for idx, item in enumerate(st.session_state.app_config_list):
        c_val = item["Valor"]
        c_hex = item.get("Color Hex", "#ffffff")

        with st.container(border=True):
            col_b1, col_b2, col_b3, col_b4 = st.columns([3, 2, 2, 1])
            with col_b1:
                st.markdown(
                    f'<div style="background-color: {c_hex}; color: #000; padding: 10px 14px; border-radius: 6px; font-weight: bold; border: 1px solid rgba(0,0,0,0.15); display: inline-block; min-width: 150px; text-align: center;">'
                    f'{c_val}</div>',
                    unsafe_allow_html=True
                )
                if c_val == st.session_state.label_ok:
                    st.caption("⭐ Cuenta como completado")
            with col_b2:
                new_hex = st.color_picker(f"Color de {c_val}", value=c_hex, key=f"cp_{idx}_{c_val}")
                if new_hex != c_hex:
                    item["Color Hex"] = new_hex
                    st.rerun()
            with col_b3:
                st.text_input("Código Hex", value=item["Color Hex"], key=f"hex_{idx}_{c_val}", disabled=True)
            with col_b4:
                if len(st.session_state.app_config_list) > 1:
                    if st.button("🗑️", key=f"del_{idx}_{c_val}", help=f"Eliminar estado {c_val}"):
                        eliminar_idx = idx

    if eliminar_idx is not None:
        st.session_state.app_config_list.pop(eliminar_idx)
        if st.session_state.label_ok not in [x["Valor"] for x in st.session_state.app_config_list]:
            if st.session_state.app_config_list:
                st.session_state.label_ok = st.session_state.app_config_list[0]["Valor"]
        st.rerun()

    st.divider()

    # Agregar nuevo estado
    with st.expander("➕ Agregar un nuevo Estado"):
        with st.form("form_nuevo_estado", clear_on_submit=True):
            col_n1, col_n2 = st.columns(2)
            with col_n1:
                nuevo_nombre = st.text_input("Nombre del nuevo Estado", placeholder="Ej: OBSERVADO, EN TRAMITE")
            with col_n2:
                nuevo_color = st.color_picker("Color para el nuevo Estado", value="#fff2cc")
            es_comp = st.checkbox("¿Este estado cuenta como expediente completado?")

            if st.form_submit_button("Agregar Estado", type="primary"):
                nombre_clean = nuevo_nombre.strip().upper()
                if not nombre_clean:
                    st.error("El nombre del estado no puede estar vacío.")
                elif nombre_clean in [x["Valor"] for x in st.session_state.app_config_list]:
                    st.error("Ese estado ya existe.")
                else:
                    st.session_state.app_config_list.append({
                        "Valor": nombre_clean,
                        "Color Hex": nuevo_color,
                        "Completado": "Sí" if es_comp else "No"
                    })
                    if es_comp:
                        st.session_state.label_ok = nombre_clean
                    st.success(f"Estado '{nombre_clean}' agregado.")
                    st.rerun()

    st.divider()
    st.subheader("Parámetros Generales de Lectura")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.selectbox(
            "¿La primera fila es encabezado?", ["Auto", "Sí", "No"],
            help="Las columnas se leen por posición: A=Expediente, B=Banco, C=Registro, D=Observaciones.",
            key="header_mode_select"
        )
    with col_p2:
        st.number_input(
            "Meta de expedientes del mes (0 = sin meta)",
            min_value=0, value=0, step=10, key="meta_input"
        )

# ---------- Instrucciones ----------
with tab_help:
    st.header("Cómo crear la key de Google (usuario virtual)")
    st.markdown("""
El usuario virtual se llama **cuenta de servicio**. Es como un correo propio de la app: solo puede
editar las hojas que tú le compartas.

**1. Crear un proyecto**
1. Entra a <https://console.cloud.google.com/> con tu cuenta de Google.
2. Arriba, en el selector de proyectos → **Proyecto nuevo** → ponle un nombre (ej. `control-expedientes`) → **Crear**.

**2. Activar las APIs**
1. Menú ☰ → **APIs y servicios → Biblioteca**.
2. Busca **Google Sheets API** → **Habilitar**.
3. Busca **Google Drive API** → **Habilitar**.

**3. Crear la cuenta de servicio**
1. Menú ☰ → **APIs y servicios → Credenciales** → **Crear credenciales → Cuenta de servicio**.
2. Nombre (ej. `editor-hojas`) → **Crear y continuar** → puedes omitir roles y usuarios → **Listo**.

**4. Descargar la key (JSON)**
1. En la lista de cuentas de servicio, abre la que creaste → pestaña **Claves**.
2. **Agregar clave → Crear clave nueva → JSON → Crear**. Se descarga un archivo `.json`.
3. ⚠️ Trátalo como una contraseña: no lo subas a GitHub ni lo compartas.

**5. Compartir la hoja con el usuario virtual**
1. Abre el archivo `.json` y copia el valor de `client_email` (termina en `iam.gserviceaccount.com`).
2. En tu hoja de Google → **Compartir** → pega ese correo → rol **Editor** → Enviar.

**6. Usarla en la app**
- **Opción A:** sube el `.json` en el sidebar o pega su contenido.
- **Opción B (recomendada):** crea `.streamlit/secrets.toml` junto a `app.py`:
""")
    st.code(
        '''[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "editor-hojas@tu-proyecto.iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"''',
        language="toml",
    )
    st.markdown("""
Copia los campos tal cual del JSON. Añade `.streamlit/secrets.toml` a tu `.gitignore`.

**Formato esperado de la hoja** (por posición de columna): A = Expediente, B = Banco,
C = Registro, D = Observaciones. Se admite una fila de encabezado o no, y columnas extra (ej. E).
Cada mes es una pestaña distinta; se elige en el sidebar.

**Problemas frecuentes**
- *SpreadsheetNotFound / 403*: falta compartir la hoja con el `client_email` como Editor.
- *APIs no habilitadas*: revisa el paso 2 (Sheets y Drive).
- *JSON inválido*: pega el contenido completo, incluyendo las llaves `{ }`.
""")
