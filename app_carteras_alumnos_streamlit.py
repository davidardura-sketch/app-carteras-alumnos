# ============================================================
# APP CARTERAS ALUMNOS - VERSIÓN 1.0 SUPABASE
# Streamlit
# ============================================================
# Objetivo:
# - Cargar una base de datos Excel con clasificación de fondos y liquidativos.
# - Calcular perfil de riesgo mediante test: Conservador, Moderado, Dinámico.
# - Definir pesos por grandes bloques: Renta Variable, Renta Fija, Liquidez.
# - Permitir filtrar fondos por categoría, subcategoría y búsqueda libre.
# - Permitir seleccionar fondos dentro de cada bloque y asignar pesos.
# - Mantener seleccionados los fondos aunque el alumno cambie los filtros.
# - Calcular backtest, liquidativo base 10.000, rentabilidad, volatilidad, Sharpe y drawdown.
# - Generar un one pager llamado "Informe de cartera".
# - Ocultar fechas reales al alumno. El eje temporal se muestra por años: 1, 2, 3, 4, 5...
#
# Ejecución local:
#   pip install streamlit pandas numpy plotly openpyxl xlsxwriter reportlab supabase
#   streamlit run app_carteras_alumnos_streamlit.py
#
# El Excel debe estar en la misma carpeta que esta app, con nombre:
#   base_datos_fondos_clasificada_con_liquidativos.xlsx
# ============================================================

import io
import json
import hashlib
import uuid
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

try:
    from supabase import create_client, Client
except ModuleNotFoundError:
    create_client = None
    Client = None


# ------------------------------------------------------------
# Configuración general
# ------------------------------------------------------------
st.set_page_config(
    page_title="Constructor de Carteras - v1.0 Supabase",
    page_icon="📊",
    layout="wide",
)

APP_DIR = Path(__file__).parent if "__file__" in globals() else Path.cwd()
DATA_DIR = APP_DIR / "data_app_carteras"
DATA_DIR.mkdir(exist_ok=True)

DEFAULT_EXCEL = APP_DIR / "base_datos_fondos_clasificada_con_liquidativos.xlsx"
SUBMISSIONS_FILE = DATA_DIR / "carteras_alumnos.csv"
PROFILE_FILE = DATA_DIR / "perfiles_riesgo_usuarios.json"
USERS_FILE = DATA_DIR / "usuarios_app.json"

RISK_FREE_RATE = 0.02
MONTHS_PER_YEAR = 12
INITIAL_NAV = 10000


# ------------------------------------------------------------
# Estilos visuales
# ------------------------------------------------------------
st.markdown(
    """
    <style>
        .main-header {
            background: linear-gradient(90deg, #0B1F3A 0%, #1F4E79 100%);
            padding: 1.4rem 1.7rem;
            border-radius: 18px;
            color: white;
            margin-bottom: 1.4rem;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.18);
        }
        .main-header h1 {
            margin-bottom: 0.25rem;
        }
        .soft-card {
            background-color: #F4F7FB;
            padding: 1rem;
            border-radius: 14px;
            border: 1px solid #D9E2EF;
        }
        .small-note {
            color: #64748B;
            font-size: 0.9rem;
        }
        .block-title {
            font-weight: 700;
            color: #0B1F3A;
            margin-top: 0.7rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Usuarios, perfiles y entregas en Supabase
# ------------------------------------------------------------
# Esta versión sustituye los JSON/CSV locales por tablas Supabase.
# La app sigue ejecutándose en Streamlit Cloud; Supabase solo almacena datos vivos.


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()


DEFAULT_USERS = {
    "profesor": {
        "password": hash_password("profesor"),
        "role": "profesor",
        "name": "Profesor",
        "internal_id": "ADMIN",
        "active": True,
    }
}


def get_supabase_credentials() -> tuple[str | None, str | None]:
    """Lee credenciales desde Streamlit Secrets.

    Formato recomendado en Streamlit Cloud / .streamlit/secrets.toml:

    [supabase]
    url = "https://xxxxxxxx.supabase.co"
    service_role_key = "eyJ..."

    También acepta claves raíz SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY.
    """
    url = None
    key = None
    try:
        if "supabase" in st.secrets:
            url = st.secrets["supabase"].get("url")
            key = st.secrets["supabase"].get("service_role_key") or st.secrets["supabase"].get("key")
        url = url or st.secrets.get("SUPABASE_URL")
        key = key or st.secrets.get("SUPABASE_SERVICE_ROLE_KEY") or st.secrets.get("SUPABASE_KEY")
    except Exception:
        pass
    return url, key


@st.cache_resource(show_spinner=False)
def get_supabase_client():
    if create_client is None:
        st.error("Falta instalar la librería supabase. Añade 'supabase' a requirements.txt.")
        st.stop()

    url, key = get_supabase_credentials()
    if not url or not key:
        st.error(
            "Faltan las credenciales de Supabase. Configura SUPABASE_URL y "
            "SUPABASE_SERVICE_ROLE_KEY en los Secrets de Streamlit Cloud."
        )
        st.stop()

    return create_client(url, key)


def safe_execute(query):
    """Ejecuta una consulta Supabase con mensaje legible si falla."""
    try:
        return query.execute()
    except Exception as e:
        st.error(f"Error de conexión/consulta con Supabase: {e}")
        st.stop()


def bootstrap_admin_if_empty() -> None:
    """Crea un profesor inicial si la tabla app_users está vacía.

    Después de entrar por primera vez, cambia esta contraseña desde el panel de profesor.
    """
    client = get_supabase_client()
    res = safe_execute(client.table("app_users").select("username").limit(1))
    if not getattr(res, "data", None):
        admin = {
            "username": "profesor",
            "password_hash": hash_password("profesor"),
            "role": "profesor",
            "name": "Profesor",
            "internal_id": "ADMIN",
            "active": True,
        }
        safe_execute(client.table("app_users").insert(admin))


def load_users() -> dict:
    bootstrap_admin_if_empty()
    client = get_supabase_client()
    res = safe_execute(client.table("app_users").select("*"))
    users = {}
    for row in res.data or []:
        users[row["username"]] = {
            "password": row.get("password_hash", ""),
            "role": row.get("role", "alumno"),
            "name": row.get("name") or row.get("username"),
            "internal_id": row.get("internal_id") or row.get("username"),
            "active": bool(row.get("active", True)),
        }
    return users


def save_users(users: dict) -> None:
    """Compatibilidad. En Supabase los cambios se hacen con funciones específicas."""
    client = get_supabase_client()
    rows = []
    for username, data in users.items():
        rows.append({
            "username": username,
            "password_hash": data.get("password", ""),
            "role": data.get("role", "alumno"),
            "name": data.get("name", username),
            "internal_id": data.get("internal_id", username),
            "active": bool(data.get("active", True)),
        })
    if rows:
        safe_execute(client.table("app_users").upsert(rows, on_conflict="username"))


def create_user(username: str, password: str, name: str, internal_id: str, role: str = "alumno") -> tuple[bool, str]:
    username = username.strip()
    if not username:
        return False, "El usuario no puede estar vacío."
    if not password:
        return False, "La contraseña no puede estar vacía."

    users = load_users()
    if username in users:
        return False, "Ese usuario ya existe."

    client = get_supabase_client()
    row = {
        "username": username,
        "password_hash": hash_password(password),
        "role": role,
        "name": name.strip() or username,
        "internal_id": internal_id.strip() or username,
        "active": True,
    }
    safe_execute(client.table("app_users").insert(row))
    return True, "Usuario creado correctamente."


def update_user_password(username: str, new_password: str) -> tuple[bool, str]:
    if not new_password:
        return False, "La nueva contraseña no puede estar vacía."
    client = get_supabase_client()
    safe_execute(
        client.table("app_users")
        .update({"password_hash": hash_password(new_password)})
        .eq("username", username)
    )
    return True, "Contraseña actualizada."


def set_user_active(username: str, active: bool) -> tuple[bool, str]:
    client = get_supabase_client()
    safe_execute(client.table("app_users").update({"active": bool(active)}).eq("username", username))
    return True, "Estado actualizado."


def delete_user_completely(username: str) -> tuple[bool, str]:
    users = load_users()
    if username not in users:
        return False, "Usuario no encontrado."
    if users[username].get("role") == "profesor":
        num_profesores = sum(1 for u in users.values() if u.get("role") == "profesor")
        if num_profesores <= 1:
            return False, "No se puede borrar el último usuario profesor."

    client = get_supabase_client()
    safe_execute(client.table("risk_profiles").delete().eq("username", username))
    safe_execute(client.table("app_users").delete().eq("username", username))
    return True, "Usuario eliminado. Las entregas históricas se conservan."


USERS = load_users()


# ------------------------------------------------------------
# Perfiles y restricciones por bloque
# ------------------------------------------------------------
PROFILE_RULES = {
    "Conservador": {
        "Renta Variable": (0, 30, 15),
        "Renta Fija": (50, 90, 70),
        "Liquidez": (5, 40, 15),
        "description": "Cartera defensiva, con prioridad en estabilidad y control de volatilidad.",
    },
    "Moderado": {
        "Renta Variable": (20, 60, 40),
        "Renta Fija": (30, 75, 50),
        "Liquidez": (0, 25, 10),
        "description": "Cartera equilibrada, combinando crecimiento y control del riesgo.",
    },
    "Dinámico": {
        "Renta Variable": (50, 90, 70),
        "Renta Fija": (5, 45, 25),
        "Liquidez": (0, 15, 5),
        "description": "Cartera orientada a crecimiento, aceptando mayor volatilidad.",
    },
}

BLOCKS = ["Renta Variable", "Renta Fija", "Liquidez"]


# ------------------------------------------------------------
# Test de perfil de riesgo
# ------------------------------------------------------------
RISK_TEST_QUESTIONS = [
    {
        "id": "1",
        "question": "Nivel de estudios",
        "options": [
            {"label": "A", "text": "Estudios universitarios", "score": 3},
            {"label": "B", "text": "Bachillerato", "score": 2},
            {"label": "C", "text": "Sin estudios", "score": 1},
        ],
    },
    {
        "id": "2a",
        "question": "¿Qué relación existe entre la rentabilidad y el riesgo?",
        "options": [
            {"label": "A", "text": "Son directamente proporcionales", "score": 1},
            {"label": "B", "text": "No existe relación alguna", "score": 0},
            {"label": "C", "text": "No sé la respuesta", "score": 0},
        ],
    },
    {
        "id": "2b",
        "question": "¿Qué significa que una inversión tenga alta liquidez?",
        "options": [
            {"label": "A", "text": "Que ofrece siempre una rentabilidad elevada", "score": 0},
            {"label": "B", "text": "Que puede convertirse fácilmente en dinero sin pérdida relevante de valor", "score": 1},
            {"label": "C", "text": "No sé la respuesta", "score": 0},
        ],
    },
    {
        "id": "2c",
        "question": "¿Qué ocurre normalmente cuando una inversión tiene más riesgo?",
        "options": [
            {"label": "A", "text": "No afecta a la rentabilidad", "score": 0},
            {"label": "B", "text": "No sé la respuesta", "score": 0},
            {"label": "C", "text": "Puede ofrecer mayor rentabilidad potencial, pero también mayores pérdidas", "score": 1},
        ],
    },
    {
        "id": "3",
        "question": "¿Cuál es su horizonte temporal de inversión?",
        "options": [
            {"label": "A", "text": "Corto plazo: entre 1 y 3 años", "score": 1},
            {"label": "B", "text": "Medio plazo: entre 3 y 5 años", "score": 2},
            {"label": "C", "text": "Largo plazo: más de 5 años", "score": 3},
        ],
    },
    {
        "id": "4",
        "question": "¿Cuál es su objetivo de inversión?",
        "options": [
            {"label": "A", "text": "Preservación de capital", "score": 1},
            {"label": "B", "text": "Aumento ligero de capital", "score": 2},
            {"label": "C", "text": "Crecimiento de capital", "score": 3},
        ],
    },
    {
        "id": "5",
        "question": "¿Qué nivel de pérdida potencial puntual estaría dispuesto a asumir?",
        "options": [
            {"label": "A", "text": "Menos del 2%", "score": 1},
            {"label": "B", "text": "Entre el 2% y el 10%", "score": 2},
            {"label": "C", "text": "Más del 10%", "score": 3},
        ],
    },
    {
        "id": "6",
        "question": "¿Cuál sería su comportamiento ante descensos fuertes de mercado?",
        "options": [
            {"label": "A", "text": "Vender", "score": 1},
            {"label": "B", "text": "Mantener", "score": 2},
            {"label": "C", "text": "Comprar", "score": 3},
        ],
    },
    {
        "id": "7",
        "question": "¿Qué porcentaje representan sus gastos periódicos sobre sus ingresos?",
        "options": [
            {"label": "A", "text": "Menos del 25%", "score": 3},
            {"label": "B", "text": "Entre el 25% y el 50%", "score": 2},
            {"label": "C", "text": "Más del 50%", "score": 1},
        ],
    },
]


def profile_from_score(score: int) -> str:
    """Mapeo del test de riesgo.

    La tabla original marca:
    - Conservador: < 9
    - Moderado: 10-17
    - Dinámico: > 18

    Para evitar huecos en 9 y 18, usamos:
    - Conservador: <= 9
    - Moderado: 10-18
    - Dinámico: >= 19
    """
    if score <= 9:
        return "Conservador"
    if score <= 18:
        return "Moderado"
    return "Dinámico"


def load_user_profiles() -> dict:
    client = get_supabase_client()
    res = safe_execute(client.table("risk_profiles").select("*"))
    profiles = {}
    for row in res.data or []:
        profiles[row["username"]] = {
            "profile": row.get("profile"),
            "score": row.get("score"),
            "timestamp": row.get("timestamp"),
            "answers": row.get("answers") or {},
        }
    return profiles


def save_user_profiles(profiles: dict) -> None:
    client = get_supabase_client()
    rows = []
    for username, data in profiles.items():
        rows.append({
            "username": username,
            "profile": data.get("profile"),
            "score": int(data.get("score", 0)) if data.get("score") not in [None, ""] else None,
            "answers": data.get("answers") or {},
            "timestamp": data.get("timestamp") or datetime.now().isoformat(timespec="seconds"),
        })
    if rows:
        safe_execute(client.table("risk_profiles").upsert(rows, on_conflict="username"))


def get_user_risk_profile(username: str) -> dict | None:
    client = get_supabase_client()
    res = safe_execute(client.table("risk_profiles").select("*").eq("username", username).limit(1))
    if not res.data:
        return None
    row = res.data[0]
    return {
        "profile": row.get("profile"),
        "score": row.get("score"),
        "timestamp": row.get("timestamp"),
        "answers": row.get("answers") or {},
    }


def save_user_risk_profile(username: str, profile_data: dict) -> None:
    client = get_supabase_client()
    row = {
        "username": username,
        "profile": profile_data.get("profile"),
        "score": int(profile_data.get("score", 0)),
        "answers": profile_data.get("answers") or {},
        "timestamp": profile_data.get("timestamp") or datetime.now().isoformat(timespec="seconds"),
    }
    safe_execute(client.table("risk_profiles").upsert(row, on_conflict="username"))


def delete_user_risk_profile(username: str) -> None:
    client = get_supabase_client()
    safe_execute(client.table("risk_profiles").delete().eq("username", username))


# ------------------------------------------------------------
# Utilidades de normalización
# ------------------------------------------------------------
def normalize_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def simplify_colname(col: str) -> str:
    return (
        str(col)
        .strip()
        .lower()
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ñ", "n")
        .replace("/", " ")
        .replace("-", " ")
        .replace("_", " ")
    )


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    simplified = {simplify_colname(c): c for c in df.columns}
    for candidate in candidates:
        cand = simplify_colname(candidate)
        for simple, original in simplified.items():
            if cand == simple or cand in simple:
                return original
    return None


def classify_block(category: str, subcategory: str = "") -> str:
    text = f"{category} {subcategory}".lower()
    if any(word in text for word in ["liquidez", "monetario", "money market", "cash"]):
        return "Liquidez"
    if any(word in text for word in ["renta variable", "equity", "acciones", "bolsa", "rv"]):
        return "Renta Variable"
    if any(word in text for word in ["renta fija", "fixed income", "bono", "bond", "credito", "crédito", "high yield", "ig", "soberana"]):
        return "Renta Fija"
    return "Otros"


def make_year_axis(df: pd.DataFrame, month_col: str = "Mes") -> tuple[list[int], list[str]]:
    """Devuelve marcas de eje X cada 12 meses: 1, 2, 3, 4, 5..."""
    max_month = int(df[month_col].max()) if not df.empty else 0
    tickvals = list(range(12, max_month + 1, 12))
    ticktext = [str(i) for i in range(1, len(tickvals) + 1)]
    return tickvals, ticktext


# ------------------------------------------------------------
# Carga del Excel
# ------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_database_from_excel(excel_bytes: bytes | None, default_path: str | None = None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Carga clasificación y liquidativos desde Excel.

    Devuelve:
    - classification: tabla maestra enriquecida con bloque.
    - navs: histórico de liquidativos con fechas internas.
    - returns: rentabilidades mensuales.
    """
    if excel_bytes is not None:
        xls = pd.ExcelFile(io.BytesIO(excel_bytes))
    elif default_path and Path(default_path).exists():
        xls = pd.ExcelFile(default_path)
    else:
        raise FileNotFoundError("No se ha encontrado el Excel de base de datos.")

    sheet_names = xls.sheet_names
    class_sheet = next((s for s in sheet_names if "clas" in simplify_colname(s)), None)
    hist_sheet = next((s for s in sheet_names if "histor" in simplify_colname(s) or "liquid" in simplify_colname(s)), None)

    if class_sheet is None:
        raise ValueError("No encuentro una hoja de clasificación en el Excel.")
    if hist_sheet is None:
        raise ValueError("No encuentro una hoja de histórico/liquidativos en el Excel.")

    classification = pd.read_excel(xls, sheet_name=class_sheet)
    navs_raw = pd.read_excel(xls, sheet_name=hist_sheet)

    # Columnas de clasificación
    id_col = find_column(classification, ["ticker", "isin", "identificador", "ticker isin", "ticker bloomberg"])
    name_col = find_column(classification, ["nombre", "fondo", "nombre del fondo", "descripcion"])
    category_col = find_column(classification, ["categoria principal", "categoria", "asset class", "clase"])
    subcategory_col = find_column(classification, ["subcategoria", "subcategoría", "tipo", "estilo"])
    region_col = find_column(classification, ["region", "región", "geografia", "geografía"])
    currency_col = find_column(classification, ["divisa", "currency"])

    if id_col is None:
        id_col = classification.columns[0]
    if name_col is None:
        name_col = classification.columns[1] if len(classification.columns) > 1 else id_col
    if category_col is None:
        category_col = classification.columns[2] if len(classification.columns) > 2 else name_col

    classification = classification.copy()
    classification["Identificador"] = classification[id_col].map(normalize_text)
    classification["Nombre"] = classification[name_col].map(normalize_text)
    classification["Categoría"] = classification[category_col].map(normalize_text)
    classification["Subcategoría"] = classification[subcategory_col].map(normalize_text) if subcategory_col else ""
    classification["Región"] = classification[region_col].map(normalize_text) if region_col else ""
    classification["Divisa"] = classification[currency_col].map(normalize_text) if currency_col else ""
    classification["Bloque"] = classification.apply(
        lambda r: classify_block(r["Categoría"], r["Subcategoría"]), axis=1
    )

    classification = classification[
        ["Identificador", "Nombre", "Categoría", "Subcategoría", "Región", "Divisa", "Bloque"]
    ].dropna(subset=["Identificador"])
    classification = classification[classification["Identificador"] != ""].drop_duplicates("Identificador")

    # Histórico de liquidativos
    navs = navs_raw.copy()
    date_col = find_column(navs, ["fecha", "date"])
    if date_col is None:
        date_col = navs.columns[0]

    navs[date_col] = pd.to_datetime(navs[date_col], errors="coerce")
    navs = navs.dropna(subset=[date_col]).set_index(date_col).sort_index()
    navs.index.name = "Fecha"

    for col in navs.columns:
        navs[col] = pd.to_numeric(navs[col], errors="coerce")

    # Intento de cruce con identificadores de la clasificación
    class_ids = set(classification["Identificador"].astype(str))
    matching_cols = [c for c in navs.columns if str(c) in class_ids]

    if len(matching_cols) == 0:
        # Fallback: intenta cruzar por columnas que contienen el identificador.
        matching_cols = []
        for col in navs.columns:
            col_text = str(col)
            if any(identifier in col_text for identifier in class_ids if identifier):
                matching_cols.append(col)

    if len(matching_cols) == 0:
        # Si no hay cruce perfecto, mantiene las columnas numéricas y avisa en pantalla luego.
        matching_cols = [c for c in navs.columns if pd.api.types.is_numeric_dtype(navs[c])]

    navs = navs[matching_cols].dropna(how="all")
    navs = navs.ffill().dropna(axis=1, how="all")
    returns = navs.pct_change().dropna(how="all")

    # Marca cobertura
    available_ids = set(str(c) for c in navs.columns)
    classification["Disponible"] = classification["Identificador"].isin(available_ids)

    return classification, navs, returns


# ------------------------------------------------------------
# Cálculos financieros
# ------------------------------------------------------------
def calculate_backtest(navs: pd.DataFrame, final_weights: dict[str, float]) -> dict:
    selected = {asset: weight for asset, weight in final_weights.items() if weight > 0}
    if not selected:
        raise ValueError("No hay activos seleccionados.")

    selected_cols = list(selected.keys())
    prices = navs[selected_cols].dropna(how="all").ffill().dropna()
    returns = prices.pct_change().dropna()

    # Quita activos sin datos suficientes
    valid_cols = [c for c in returns.columns if returns[c].dropna().shape[0] > 6]
    returns = returns[valid_cols].dropna()

    if len(valid_cols) == 0:
        raise ValueError("Los fondos seleccionados no tienen datos suficientes para calcular el backtest.")

    weights = np.array([selected[c] / 100 for c in valid_cols], dtype=float)
    weights = weights / weights.sum()

    portfolio_returns = returns.dot(weights)
    liquidativo = INITIAL_NAV * (1 + portfolio_returns).cumprod()

    total_return = liquidativo.iloc[-1] / liquidativo.iloc[0] - 1
    years = len(portfolio_returns) / MONTHS_PER_YEAR
    annual_return = (1 + total_return) ** (1 / years) - 1 if years > 0 else np.nan
    annual_volatility = portfolio_returns.std() * np.sqrt(MONTHS_PER_YEAR)
    sharpe = (annual_return - RISK_FREE_RATE) / annual_volatility if annual_volatility and annual_volatility > 0 else np.nan

    rolling_max = liquidativo.cummax()
    drawdown = liquidativo / rolling_max - 1
    max_drawdown = drawdown.min()

    # Métricas individuales de los fondos seleccionados en el mismo periodo analizado.
    fund_metrics_rows = []
    fund_weights = dict(zip(valid_cols, weights * 100))
    fund_navs = pd.DataFrame(index=returns.index)
    for col in valid_cols:
        fund_returns = returns[col].dropna()
        if fund_returns.empty:
            continue
        fund_nav = INITIAL_NAV * (1 + fund_returns).cumprod()
        fund_navs[col] = fund_nav
        fund_total_return = fund_nav.iloc[-1] / fund_nav.iloc[0] - 1
        fund_years = len(fund_returns) / MONTHS_PER_YEAR
        fund_annual_return = (1 + fund_total_return) ** (1 / fund_years) - 1 if fund_years > 0 else np.nan
        fund_annual_volatility = fund_returns.std() * np.sqrt(MONTHS_PER_YEAR)
        fund_metrics_rows.append({
            "Identificador": col,
            "Peso final": fund_weights.get(col, np.nan),
            "Rentabilidad acumulada": fund_total_return,
            "Rentabilidad anualizada": fund_annual_return,
            "Volatilidad anualizada": fund_annual_volatility,
        })

    fund_metrics = pd.DataFrame(fund_metrics_rows)

    return {
        "selected_assets": valid_cols,
        "weights": fund_weights,
        "portfolio_returns": portfolio_returns,
        "liquidativo": liquidativo,
        "total_return": total_return,
        "annual_return": annual_return,
        "annual_volatility": annual_volatility,
        "sharpe": sharpe,
        "drawdown": drawdown,
        "max_drawdown": max_drawdown,
        "fund_metrics": fund_metrics,
        "fund_navs": fund_navs,
    }


def format_pct(x: float) -> str:
    if pd.isna(x):
        return "n.d."
    return f"{x:.2%}"


def format_num(x: float) -> str:
    if pd.isna(x):
        return "n.d."
    return f"{x:.2f}"


def prepare_fund_metrics_table(result: dict, portfolio_detail: pd.DataFrame) -> pd.DataFrame:
    """Une métricas individuales calculadas con los nombres y categorías de la cartera."""
    fund_metrics = result.get("fund_metrics")
    if fund_metrics is None or fund_metrics.empty:
        return pd.DataFrame()

    metrics = fund_metrics.copy()
    if portfolio_detail is None or portfolio_detail.empty:
        return metrics

    cols = [
        "Identificador",
        "Nombre",
        "Bloque",
        "Categoría",
        "Subcategoría",
        "Peso final",
    ]
    available_cols = [c for c in cols if c in portfolio_detail.columns]
    info = portfolio_detail[available_cols].copy()
    info["Identificador"] = info["Identificador"].astype(str)
    metrics["Identificador"] = metrics["Identificador"].astype(str)

    # Si existe Peso final en ambas fuentes, priorizamos el peso de la cartera visual.
    metrics = metrics.drop(columns=["Peso final"], errors="ignore")
    merged = metrics.merge(info, on="Identificador", how="left")

    ordered_cols = [
        "Identificador",
        "Nombre",
        "Bloque",
        "Categoría",
        "Subcategoría",
        "Peso final",
        "Rentabilidad acumulada",
        "Rentabilidad anualizada",
        "Volatilidad anualizada",
    ]
    ordered_cols = [c for c in ordered_cols if c in merged.columns]
    return merged[ordered_cols].sort_values("Peso final", ascending=False) if "Peso final" in merged.columns else merged[ordered_cols]


def format_fund_metrics_for_display(fund_metrics: pd.DataFrame) -> pd.DataFrame:
    display = fund_metrics.copy()
    if "Peso final" in display.columns:
        display["Peso final"] = pd.to_numeric(display["Peso final"], errors="coerce").map(lambda x: f"{x:.2f}%" if pd.notna(x) else "")
    for col in ["Rentabilidad acumulada", "Rentabilidad anualizada", "Volatilidad anualizada"]:
        if col in display.columns:
            display[col] = pd.to_numeric(display[col], errors="coerce").map(lambda x: f"{x:.2%}" if pd.notna(x) else "")
    return display


# ------------------------------------------------------------
# One pager Excel
# ------------------------------------------------------------
def create_one_pager_excel(result: dict, final_weights_df: pd.DataFrame, profile: str) -> bytes:
    output = io.BytesIO()

    liquidativo_df = pd.DataFrame(
        {
            "Mes": np.arange(1, len(result["liquidativo"]) + 1),
            "Año": np.ceil(np.arange(1, len(result["liquidativo"]) + 1) / 12).astype(int),
            "Liquidativo": result["liquidativo"].values,
        }
    )

    drawdown_df = pd.DataFrame(
        {
            "Mes": np.arange(1, len(result["drawdown"]) + 1),
            "Año": np.ceil(np.arange(1, len(result["drawdown"]) + 1) / 12).astype(int),
            "Drawdown": result["drawdown"].values,
        }
    )

    metrics_df = pd.DataFrame(
        [
            ["Perfil", profile],
            ["Base inicial del liquidativo", INITIAL_NAV],
            ["Rentabilidad acumulada", result["total_return"]],
            ["Rentabilidad anualizada", result["annual_return"]],
            ["Volatilidad anualizada", result["annual_volatility"]],
            ["Ratio Sharpe", result["sharpe"]],
            ["Máximo drawdown", result["max_drawdown"]],
            ["Activo libre de riesgo usado", RISK_FREE_RATE],
        ],
        columns=["Métrica", "Valor"],
    )

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        sheet_name = "Informe de cartera"
        ws = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = ws

        fmt_title = workbook.add_format({"bold": True, "font_size": 20, "font_color": "white", "bg_color": "#0B1F3A"})
        fmt_subtitle = workbook.add_format({"font_size": 11, "font_color": "#475569"})
        fmt_header = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E79", "border": 1})
        fmt_text = workbook.add_format({"border": 1})
        fmt_pct = workbook.add_format({"num_format": "0.00%", "border": 1})
        fmt_num = workbook.add_format({"num_format": "0.00", "border": 1})
        fmt_nav = workbook.add_format({"num_format": "#,##0", "border": 1})
        fmt_note = workbook.add_format({"italic": True, "font_color": "#64748B"})

        ws.merge_range("A1:H2", "Informe de cartera", fmt_title)
        ws.write("A3", f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", fmt_subtitle)
        ws.write("A4", "Las fechas del periodo histórico se ocultan deliberadamente en el ejercicio.", fmt_note)

        # Métricas
        ws.write("A6", "Métrica", fmt_header)
        ws.write("B6", "Valor", fmt_header)
        for i, row in metrics_df.iterrows():
            excel_row = 6 + i
            ws.write(excel_row, 0, row["Métrica"], fmt_text)
            value = row["Valor"]
            if row["Métrica"] in ["Rentabilidad acumulada", "Rentabilidad anualizada", "Volatilidad anualizada", "Máximo drawdown", "Activo libre de riesgo usado"]:
                ws.write_number(excel_row, 1, float(value), fmt_pct)
            elif row["Métrica"] in ["Ratio Sharpe"]:
                ws.write_number(excel_row, 1, float(value), fmt_num)
            elif row["Métrica"] in ["Base inicial del liquidativo"]:
                ws.write_number(excel_row, 1, float(value), fmt_nav)
            else:
                ws.write(excel_row, 1, str(value), fmt_text)

        # Pesos
        start_row = 6
        start_col = 3
        ws.write(start_row, start_col, "Fondo", fmt_header)
        ws.write(start_row, start_col + 1, "Bloque", fmt_header)
        ws.write(start_row, start_col + 2, "Peso", fmt_header)
        for i, row in final_weights_df.iterrows():
            r = start_row + 1 + i
            ws.write(r, start_col, row["Nombre"], fmt_text)
            ws.write(r, start_col + 1, row["Bloque"], fmt_text)
            ws.write_number(r, start_col + 2, row["Peso final"] / 100, fmt_pct)

        # Datos liquidativo ocultos a la derecha
        data_start_row = 22
        ws.write(data_start_row, 0, "Mes", fmt_header)
        ws.write(data_start_row, 1, "Año", fmt_header)
        ws.write(data_start_row, 2, "Liquidativo", fmt_header)
        for i, row in liquidativo_df.iterrows():
            r = data_start_row + 1 + i
            ws.write_number(r, 0, int(row["Mes"]))
            ws.write_number(r, 1, int(row["Año"]))
            ws.write_number(r, 2, float(row["Liquidativo"]))

        chart = workbook.add_chart({"type": "line"})
        chart.add_series({
            "name": "Liquidativo",
            "categories": [sheet_name, data_start_row + 1, 0, data_start_row + len(liquidativo_df), 0],
            "values": [sheet_name, data_start_row + 1, 2, data_start_row + len(liquidativo_df), 2],
            "line": {"color": "#1F4E79", "width": 2.5},
        })
        chart.set_title({"name": "Evolución del liquidativo base 10.000"})
        chart.set_x_axis({"name": "Año de simulación", "major_unit": 12})
        chart.set_y_axis({"name": "Liquidativo"})
        chart.set_legend({"none": True})
        chart.set_size({"width": 720, "height": 330})
        ws.insert_chart("A17", chart)

        ws.set_column("A:A", 28)
        ws.set_column("B:C", 18)
        ws.set_column("D:D", 45)
        ws.set_column("E:E", 18)
        ws.set_column("F:F", 14)

        # Hojas auxiliares
        metrics_df.to_excel(writer, sheet_name="Métricas", index=False)
        final_weights_df.to_excel(writer, sheet_name="Composición", index=False)
        fund_metrics_df = prepare_fund_metrics_table(result, final_weights_df)
        if not fund_metrics_df.empty:
            fund_metrics_df.to_excel(writer, sheet_name="Métricas fondos", index=False)
        liquidativo_df.to_excel(writer, sheet_name="Liquidativo", index=False)
        drawdown_df.to_excel(writer, sheet_name="Drawdown", index=False)

    output.seek(0)
    return output.read()



# ------------------------------------------------------------
# Informe PDF para el profesor
# ------------------------------------------------------------
def create_portfolio_pdf_report(selected_row: pd.Series, portfolio_detail: pd.DataFrame, result: dict) -> bytes:
    """Genera un PDF completo de la cartera del alumno.

    Incluye métricas, gráfico de liquidativo, drawdown y cartera completa.
    Las fechas reales no aparecen en el informe: se muestran meses y años de simulación.
    """
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics.charts.lineplots import LinePlot

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.4 * cm,
        leftMargin=1.4 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Title"],
        fontSize=20,
        leading=24,
        textColor=colors.HexColor("#0B1F3A"),
        alignment=TA_CENTER,
        spaceAfter=12,
    )
    section_style = ParagraphStyle(
        "SectionTitle",
        parent=styles["Heading2"],
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#1F4E79"),
        spaceBefore=10,
        spaceAfter=6,
    )
    normal_style = ParagraphStyle(
        "CustomNormal",
        parent=styles["BodyText"],
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#334155"),
        alignment=TA_LEFT,
    )
    small_style = ParagraphStyle(
        "Small",
        parent=styles["BodyText"],
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#64748B"),
    )

    def _pct(x):
        return format_pct(float(x)) if pd.notna(x) else "n.d."

    def _num(x):
        return format_num(float(x)) if pd.notna(x) else "n.d."

    story = []
    story.append(Paragraph("Informe completo de cartera", title_style))
    student_name = selected_row.get("name", "")
    internal_id = selected_row.get("internal_id", "")
    timestamp = selected_row.get("timestamp", "")
    profile = selected_row.get("profile", "")
    story.append(Paragraph(f"Alumno: <b>{student_name}</b> &nbsp;&nbsp; ID interno: <b>{internal_id}</b>", normal_style))
    story.append(Paragraph(f"Perfil de riesgo: <b>{profile}</b> &nbsp;&nbsp; Fecha de entrega: <b>{timestamp}</b>", normal_style))
    story.append(Paragraph("Las fechas reales del histórico permanecen ocultas. El eje temporal se expresa en meses y años de simulación.", small_style))
    story.append(Spacer(1, 10))

    metrics_data = [
        ["Métrica", "Valor"],
        ["Base inicial del liquidativo", f"{INITIAL_NAV:,.0f}"],
        ["Rentabilidad acumulada", _pct(result["total_return"])],
        ["Rentabilidad anualizada", _pct(result["annual_return"])],
        ["Volatilidad anualizada", _pct(result["annual_volatility"])],
        ["Ratio Sharpe", _num(result["sharpe"])],
        ["Máximo drawdown", _pct(result["max_drawdown"])],
        ["Activo libre de riesgo usado", _pct(RISK_FREE_RATE)],
    ]
    metrics_table = Table(metrics_data, colWidths=[7.2 * cm, 4.0 * cm])
    metrics_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B1F3A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
    ]))
    story.append(Paragraph("Resumen de resultados", section_style))
    story.append(metrics_table)
    story.append(Spacer(1, 12))

    def make_line_chart(series, title, y_min=None, y_max=None, is_pct=False):
        values = [float(v) for v in series.values]
        n = len(values)
        drawing = Drawing(500, 230)
        chart = LinePlot()
        chart.x = 45
        chart.y = 35
        chart.height = 160
        chart.width = 420
        chart.data = [[(i + 1, values[i]) for i in range(n)]]
        chart.lines[0].strokeColor = colors.HexColor("#1F4E79")
        chart.lines[0].strokeWidth = 2
        chart.xValueAxis.valueMin = 1
        chart.xValueAxis.valueMax = max(12, int(np.ceil(n / 12) * 12))
        chart.xValueAxis.valueStep = 12
        chart.xValueAxis.labelTextFormat = lambda x: str(int(np.ceil(x / 12))) if x >= 1 else ""
        chart.xValueAxis.labels.fontSize = 7
        chart.yValueAxis.labels.fontSize = 7
        if y_min is not None:
            chart.yValueAxis.valueMin = y_min
        if y_max is not None:
            chart.yValueAxis.valueMax = y_max
        if is_pct:
            chart.yValueAxis.labelTextFormat = lambda y: f"{y:.0%}"
        drawing.add(chart)
        from reportlab.graphics.shapes import String
        drawing.add(String(250, 210, title, textAnchor="middle", fontSize=11, fillColor=colors.HexColor("#0B1F3A")))
        drawing.add(String(250, 8, "Año de simulación", textAnchor="middle", fontSize=8, fillColor=colors.HexColor("#64748B")))
        return drawing

    story.append(Paragraph("Evolución del liquidativo", section_style))
    liq = result["liquidativo"]
    y_min = max(0, float(liq.min()) * 0.95)
    y_max = float(liq.max()) * 1.05
    story.append(make_line_chart(liq, "Liquidativo base 10.000", y_min=y_min, y_max=y_max))
    story.append(Spacer(1, 8))

    story.append(Paragraph("Drawdown", section_style))
    dd = result["drawdown"]
    story.append(make_line_chart(dd, "Drawdown de la cartera", y_min=min(float(dd.min()) * 1.15, -0.01), y_max=0.01, is_pct=True))
    fund_metrics_table_df = prepare_fund_metrics_table(result, portfolio_detail)
    if not fund_metrics_table_df.empty:
        story.append(Paragraph("Comportamiento de los fondos en el periodo analizado", section_style))
        story.append(Paragraph("Las métricas se calculan sobre el mismo periodo oculto usado para el backtest de la cartera.", small_style))
        fm_rows = [["Fondo", "Peso", "Rent. acum.", "Rent. anual.", "Vol. anual."]]
        for _, r in fund_metrics_table_df.iterrows():
            fm_rows.append([
                Paragraph(str(r.get("Nombre", r.get("Identificador", "")))[:90], small_style),
                f"{float(r.get('Peso final', 0.0)):.2f}%" if pd.notna(r.get("Peso final", np.nan)) else "",
                _pct(r.get("Rentabilidad acumulada", np.nan)),
                _pct(r.get("Rentabilidad anualizada", np.nan)),
                _pct(r.get("Volatilidad anualizada", np.nan)),
            ])
        fm_table = Table(fm_rows, colWidths=[7.2 * cm, 2.0 * cm, 2.4 * cm, 2.4 * cm, 2.4 * cm], repeatRows=1)
        fm_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B1F3A")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.3),
        ]))
        story.append(fm_table)

    story.append(PageBreak())

    story.append(Paragraph("Cartera completa", section_style))
    portfolio = portfolio_detail.copy()
    portfolio["Peso final"] = pd.to_numeric(portfolio["Peso final"], errors="coerce").fillna(0.0)
    portfolio = portfolio.sort_values("Peso final", ascending=False)

    table_rows = [["Fondo", "Bloque", "Categoría", "Subcategoría", "Peso"]]
    for _, r in portfolio.iterrows():
        table_rows.append([
            Paragraph(str(r.get("Nombre", ""))[:80], small_style),
            Paragraph(str(r.get("Bloque", ""))[:30], small_style),
            Paragraph(str(r.get("Categoría", ""))[:35], small_style),
            Paragraph(str(r.get("Subcategoría", ""))[:35], small_style),
            f"{float(r.get('Peso final', 0.0)):.2f}%",
        ])

    portfolio_table = Table(table_rows, colWidths=[6.1 * cm, 2.4 * cm, 3.0 * cm, 3.0 * cm, 1.8 * cm], repeatRows=1)
    portfolio_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B1F3A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("ALIGN", (-1, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
    ]))
    story.append(portfolio_table)
    story.append(Spacer(1, 12))

    story.append(Paragraph("Resumen por bloques", section_style))
    if "Bloque" in portfolio.columns:
        by_block = portfolio.groupby("Bloque", dropna=False)["Peso final"].sum().reset_index().sort_values("Peso final", ascending=False)
        block_rows = [["Bloque", "Peso"]] + [[str(r["Bloque"]), f"{float(r['Peso final']):.2f}%"] for _, r in by_block.iterrows()]
        block_table = Table(block_rows, colWidths=[8 * cm, 3 * cm])
        block_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
            ("ALIGN", (1, 1), (1, -1), "RIGHT"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]))
        story.append(block_table)

    story.append(Paragraph("Resumen por categorías", section_style))
    if "Categoría" in portfolio.columns:
        by_cat = portfolio.groupby("Categoría", dropna=False)["Peso final"].sum().reset_index().sort_values("Peso final", ascending=False)
        cat_rows = [["Categoría", "Peso"]] + [[str(r["Categoría"]), f"{float(r['Peso final']):.2f}%"] for _, r in by_cat.iterrows()]
        cat_table = Table(cat_rows, colWidths=[8 * cm, 3 * cm])
        cat_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
            ("ALIGN", (1, 1), (1, -1), "RIGHT"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]))
        story.append(cat_table)

    doc.build(story)
    buffer.seek(0)
    return buffer.read()

# ------------------------------------------------------------
# Guardado de entregas
# ------------------------------------------------------------
def save_submission(username: str, user_info: dict, profile: str, final_weights_df: pd.DataFrame, result: dict):
    client = get_supabase_client()
    submission_id = str(uuid.uuid4())
    portfolio_records = final_weights_df.copy()
    portfolio_records["Peso final"] = pd.to_numeric(portfolio_records["Peso final"], errors="coerce").fillna(0.0)
    portfolio_json = portfolio_records.to_dict(orient="records")
    portfolio_summary = "; ".join(
        f"{r['Nombre']} ({r['Peso final']:.2f}%)"
        for _, r in portfolio_records.sort_values("Peso final", ascending=False).iterrows()
    )

    submission_row = {
        "id": submission_id,
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "username": username,
        "internal_id": user_info["internal_id"],
        "name": user_info["name"],
        "profile": profile,
        "initial_nav": float(INITIAL_NAV),
        "total_return": float(result["total_return"]),
        "annual_return": float(result["annual_return"]),
        "annual_volatility": float(result["annual_volatility"]),
        "sharpe": float(result["sharpe"]),
        "max_drawdown": float(result["max_drawdown"]),
        "num_fondos": int(len(portfolio_records)),
        "portfolio_summary": portfolio_summary,
    }

    position_rows = []
    for _, r in portfolio_records.iterrows():
        position_rows.append({
            "submission_id": submission_id,
            "identifier": str(r.get("Identificador", "")),
            "fund_name": str(r.get("Nombre", "")),
            "block": str(r.get("Bloque", "")),
            "category": str(r.get("Categoría", "")),
            "subcategory": str(r.get("Subcategoría", "")),
            "internal_weight": float(r.get("Peso interno", 0) or 0),
            "final_weight": float(r.get("Peso final", 0) or 0),
        })

    safe_execute(client.table("submissions").insert(submission_row))
    if position_rows:
        safe_execute(client.table("submission_positions").insert(position_rows))


def load_submissions() -> pd.DataFrame:
    """Carga entregas desde Supabase y reconstruye portfolio_json para reutilizar la vista profesor."""
    client = get_supabase_client()
    subs_res = safe_execute(client.table("submissions").select("*").order("timestamp", desc=True))
    submissions = subs_res.data or []
    if not submissions:
        return pd.DataFrame()

    ids = [row["id"] for row in submissions if row.get("id")]
    positions_by_submission = {sid: [] for sid in ids}
    if ids:
        pos_res = safe_execute(client.table("submission_positions").select("*").in_("submission_id", ids))
        for pos in pos_res.data or []:
            record = {
                "Identificador": pos.get("identifier"),
                "Nombre": pos.get("fund_name"),
                "Bloque": pos.get("block"),
                "Categoría": pos.get("category"),
                "Subcategoría": pos.get("subcategory"),
                "Peso interno": pos.get("internal_weight"),
                "Peso final": pos.get("final_weight"),
            }
            positions_by_submission.setdefault(pos.get("submission_id"), []).append(record)

    rows = []
    for row in submissions:
        out = dict(row)
        records = positions_by_submission.get(row.get("id"), [])
        out["portfolio_json"] = json.dumps(records, ensure_ascii=False)
        for rec in records:
            out[f"peso_{rec.get('Identificador')}"] = rec.get("Peso final")
        rows.append(out)
    return pd.DataFrame(rows)


def portfolio_from_submission_row(row: pd.Series, classification: pd.DataFrame | None = None) -> pd.DataFrame:
    """Reconstruye la cartera completa de una entrega.

    Prioridad:
    1) portfolio_json, disponible en entregas nuevas.
    2) columnas peso_<Identificador>, compatible con entregas anteriores.
    """
    if "portfolio_json" in row.index and pd.notna(row.get("portfolio_json")) and str(row.get("portfolio_json")).strip():
        try:
            records = json.loads(row.get("portfolio_json"))
            portfolio = pd.DataFrame(records)
            if not portfolio.empty:
                return portfolio
        except Exception:
            pass

    # Compatibilidad con entregas antiguas: reconstruye desde columnas peso_*.
    records = []
    class_lookup = {}
    if classification is not None and not classification.empty:
        class_lookup = classification.set_index("Identificador").to_dict(orient="index")

    for col, value in row.items():
        if not str(col).startswith("peso_"):
            continue
        if pd.isna(value):
            continue
        try:
            weight = float(value)
        except Exception:
            continue
        if abs(weight) < 1e-9:
            continue

        identifier = str(col).replace("peso_", "", 1)
        meta = class_lookup.get(identifier, {})
        records.append(
            {
                "Identificador": identifier,
                "Nombre": meta.get("Nombre", identifier),
                "Bloque": meta.get("Bloque", ""),
                "Categoría": meta.get("Categoría", ""),
                "Subcategoría": meta.get("Subcategoría", ""),
                "Peso final": weight,
            }
        )

    return pd.DataFrame(records)


# ------------------------------------------------------------
# Login
# ------------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "username" not in st.session_state:
    st.session_state.username = None


def login_page():
    st.markdown(
        """
        <div class="main-header">
            <h1>Constructor de Carteras</h1>
            <p>Acceso a la herramienta de construcción y análisis de carteras.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("Acceso")
        username = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        if st.button("Entrar", type="primary"):
            users = load_users()
            if (
                username in users
                and users[username].get("password") == hash_password(password)
                and users[username].get("active", True)
            ):
                st.session_state.authenticated = True
                st.session_state.username = username
                st.rerun()
            elif username in users and not users[username].get("active", True):
                st.error("Usuario desactivado. Contacta con el profesor.")
            else:
                st.error("Usuario o contraseña incorrectos.")

    with col2:
        st.markdown(
            """
            <div class="soft-card">
                <h3>Bienvenido</h3>
                <p>Introduce tus credenciales para acceder a la herramienta.</p>
                <p class="small-note">Si no tienes usuario o has olvidado la contraseña, contacta con el profesor.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )




# ------------------------------------------------------------
# Vista test de perfil de riesgo
# ------------------------------------------------------------
def risk_profile_test_page(username: str, user_info: dict):
    st.markdown(
        f"""
        <div class="main-header">
            <h1>Test de perfil de riesgo</h1>
            <p>{user_info['name']} · Antes de construir la cartera, completa el test una sola vez.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.info(
        "El perfil de riesgo no se elige manualmente. Se calcula a partir de este test y queda guardado en tu usuario para próximas sesiones."
    )

    with st.form("risk_test_form"):
        answers = {}
        for q in RISK_TEST_QUESTIONS:
            st.markdown(f"**{q['id']}. {q['question']}**")
            option_labels = [
                f"{opt['label']} · {opt['text']}"
                for opt in q["options"]
            ]
            selected = st.radio(
                label="Selecciona una respuesta",
                options=option_labels,
                key=f"risk_q_{q['id']}",
                label_visibility="collapsed",
            )
            selected_label = selected.split(" · ")[0]
            selected_option = next(opt for opt in q["options"] if opt["label"] == selected_label)
            answers[q["id"]] = {
                "label": selected_option["label"],
                "text": selected_option["text"],
                "score": selected_option["score"],
            }
            st.write("")

        submitted = st.form_submit_button("Guardar test y continuar", type="primary")

    if submitted:
        total_score = int(sum(answer["score"] for answer in answers.values()))
        profile = profile_from_score(total_score)
        profile_data = {
            "profile": profile,
            "score": total_score,
            "answers": answers,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        }
        save_user_risk_profile(username, profile_data)

        # Limpiamos estado de cartera previo para que el perfil nuevo no arrastre selecciones anteriores.
        for key in list(st.session_state.keys()):
            if key.startswith("selected_ids_") or key.startswith("weight_") or key in ["result", "final_weights_df", "profile"]:
                del st.session_state[key]

        st.success(f"Test guardado. Tu perfil asignado es: {profile} · Puntuación: {total_score}")
        st.rerun()


# ------------------------------------------------------------
# Vista alumno
# ------------------------------------------------------------
def student_page(username: str, user_info: dict):
    st.markdown(
        f"""
        <div class="main-header">
            <h1>Construcción de cartera</h1>
            <p>{user_info['name']} · Perfil asignado por test, pesos por bloque y fondos concretos.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.warning("Las fechas históricas están ocultas deliberadamente. El gráfico mostrará años de simulación, no fechas reales.")

    with st.sidebar:
        st.subheader("Base de datos")
        uploaded = st.file_uploader("Cargar Excel de fondos", type=["xlsx"])
        st.caption("Si no cargas archivo, la app intentará usar el Excel de la misma carpeta.")

    try:
        excel_bytes = uploaded.read() if uploaded is not None else None
        classification, navs, returns = load_database_from_excel(excel_bytes, str(DEFAULT_EXCEL))
    except Exception as e:
        st.error(f"No se ha podido cargar la base de datos: {e}")
        st.stop()

    available = classification[classification["Disponible"]].copy()
    available = available[available["Bloque"].isin(BLOCKS)].copy()

    if available.empty:
        st.error("No hay fondos disponibles cruzados entre clasificación e histórico.")
        st.stop()

    profile_data = get_user_risk_profile(username)
    if profile_data is None:
        st.warning("Todavía no tienes un perfil de riesgo asignado. Completa primero el test.")
        risk_profile_test_page(username, user_info)
        st.stop()

    profile = profile_data["profile"]
    score = profile_data.get("score", "n.d.")
    rules = PROFILE_RULES[profile]

    st.subheader("1. Perfil de riesgo asignado")
    col_profile, col_score = st.columns([2, 1])
    col_profile.metric("Perfil", profile)
    col_score.metric("Puntuación test", score)
    st.info(rules["description"])

    st.subheader("2. Asignación por grandes bloques")
    st.caption("Define cuánto peso quieres tener en Renta Variable, Renta Fija y Liquidez. La suma debe ser 100%.")

    block_weights = {}
    block_cols = st.columns(3)
    for i, block in enumerate(BLOCKS):
        min_w, max_w, default_w = rules[block]
        with block_cols[i]:
            block_weights[block] = st.slider(
                block,
                min_value=int(min_w),
                max_value=int(max_w),
                value=int(default_w),
                step=5,
            )

    total_block_weight = sum(block_weights.values())
    st.metric("Suma de bloques", f"{total_block_weight:.0f}%")

    if total_block_weight != 100:
        st.error("La suma de Renta Variable, Renta Fija y Liquidez debe ser exactamente 100%.")
        st.stop()

    st.subheader("3. Selección de fondos dentro de cada bloque")
    st.caption("Dentro de cada categoría, distribuye el 100% interno entre los fondos que quieras usar.")

    final_rows = []
    internal_errors = []

    for block in BLOCKS:
        target_weight = block_weights[block]
        if target_weight == 0:
            continue

        st.markdown(f"<p class='block-title'>{block} · peso total en cartera: {target_weight:.0f}%</p>", unsafe_allow_html=True)
        funds_block = available[available["Bloque"] == block].copy()

        if funds_block.empty:
            internal_errors.append(f"No hay fondos disponibles para {block}.")
            continue

        # ----------------------------------------------------
        # Filtros por categoría, subcategoría y búsqueda libre
        # ----------------------------------------------------
        filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 1.5])

        with filter_col1:
            available_categories = sorted([
                c for c in funds_block["Categoría"].dropna().unique().tolist()
                if str(c).strip() != ""
            ])
            selected_category = st.selectbox(
                f"Filtrar categoría en {block}",
                options=["Todas"] + available_categories,
                key=f"cat_filter_{block}",
            )

        filtered_block = funds_block.copy()
        if selected_category != "Todas":
            filtered_block = filtered_block[filtered_block["Categoría"] == selected_category]

        with filter_col2:
            available_subcategories = sorted([
                c for c in filtered_block["Subcategoría"].dropna().unique().tolist()
                if str(c).strip() != ""
            ])
            selected_subcategory = st.selectbox(
                f"Filtrar subcategoría en {block}",
                options=["Todas"] + available_subcategories,
                key=f"subcat_filter_{block}",
            )

        if selected_subcategory != "Todas":
            filtered_block = filtered_block[filtered_block["Subcategoría"] == selected_subcategory]

        with filter_col3:
            search_text = st.text_input(
                f"Buscar fondo en {block}",
                placeholder="Escribe parte del nombre o identificador...",
                key=f"search_{block}",
            )

        if search_text.strip():
            search_norm = search_text.strip().lower()
            filtered_block = filtered_block[
                filtered_block["Nombre"].str.lower().str.contains(search_norm, na=False)
                | filtered_block["Identificador"].str.lower().str.contains(search_norm, na=False)
                | filtered_block["Categoría"].str.lower().str.contains(search_norm, na=False)
                | filtered_block["Subcategoría"].str.lower().str.contains(search_norm, na=False)
            ]

        st.caption(f"Fondos disponibles tras filtros: {len(filtered_block)}")

        # ----------------------------------------------------
        # Selección persistente de fondos — versión robusta
        # ----------------------------------------------------
        # Bug resuelto:
        # No usamos un multiselect para mantener la lista de fondos seleccionados,
        # porque al cambiar los filtros Streamlit puede recalcular opciones y borrar
        # selecciones previas. Ahora guardamos IDs en session_state y la eliminación
        # se hace explícitamente con un botón.
        selected_ids_key = f"selected_ids_{block}"
        if selected_ids_key not in st.session_state:
            st.session_state[selected_ids_key] = []

        all_label_to_id = {
            f"{row.Nombre} | {row.Identificador}": row.Identificador
            for row in funds_block.itertuples()
        }
        all_id_to_label = {v: k for k, v in all_label_to_id.items()}

        # Limpieza de ids antiguos o no disponibles
        current_available_ids = set(funds_block["Identificador"])
        st.session_state[selected_ids_key] = [
            fund_id for fund_id in st.session_state[selected_ids_key]
            if fund_id in current_available_ids
        ]

        filtered_label_to_id = {
            f"{row.Nombre} | {row.Identificador}": row.Identificador
            for row in filtered_block.itertuples()
            if row.Identificador not in st.session_state[selected_ids_key]
        }

        if filtered_label_to_id:
            add_col1, add_col2 = st.columns([3, 1])
            with add_col1:
                label_to_add = st.selectbox(
                    f"Fondo para añadir a {block}",
                    options=[""] + list(filtered_label_to_id.keys()),
                    format_func=lambda x: "Selecciona un fondo..." if x == "" else x,
                    key=f"add_one_{block}",
                    help="Puedes cambiar los filtros y seguir añadiendo fondos. Los ya seleccionados se mantienen.",
                )
            with add_col2:
                st.write("")
                st.write("")
                if st.button(f"Añadir a {block}", key=f"add_button_{block}"):
                    if label_to_add:
                        fund_id = filtered_label_to_id[label_to_add]
                        if fund_id not in st.session_state[selected_ids_key]:
                            st.session_state[selected_ids_key].append(fund_id)
                        # Si había resultados calculados, los invalidamos porque la cartera ha cambiado.
                        for k in ["result", "final_weights_df", "profile"]:
                            if k in st.session_state:
                                del st.session_state[k]
                        st.rerun()
        else:
            st.info(f"No hay fondos nuevos que añadir a {block} con los filtros actuales.")

        selected_ids = st.session_state[selected_ids_key]

        if selected_ids:
            selected_table = funds_block[funds_block["Identificador"].isin(selected_ids)][
                ["Nombre", "Identificador", "Categoría", "Subcategoría"]
            ].copy()
            selected_table.insert(0, "Nº", range(1, len(selected_table) + 1))
            st.markdown(f"**Fondos seleccionados en {block}**")
            st.dataframe(selected_table, use_container_width=True, hide_index=True)

            remove_col1, remove_col2 = st.columns([3, 1])
            selected_labels_all = [all_id_to_label[fund_id] for fund_id in selected_ids if fund_id in all_id_to_label]
            with remove_col1:
                label_to_remove = st.selectbox(
                    f"Eliminar fondo de {block}",
                    options=[""] + selected_labels_all,
                    format_func=lambda x: "No eliminar ningún fondo" if x == "" else x,
                    key=f"remove_one_{block}",
                )
            with remove_col2:
                st.write("")
                st.write("")
                if st.button(f"Eliminar de {block}", key=f"remove_button_{block}"):
                    if label_to_remove:
                        remove_id = all_label_to_id[label_to_remove]
                        st.session_state[selected_ids_key] = [
                            fund_id for fund_id in selected_ids if fund_id != remove_id
                        ]
                        weight_key = f"weight_{block}_{remove_id}"
                        if weight_key in st.session_state:
                            del st.session_state[weight_key]
                        for k in ["result", "final_weights_df", "profile"]:
                            if k in st.session_state:
                                del st.session_state[k]
                        st.rerun()
        else:
            st.info(f"Todavía no has añadido fondos de {block}.")

        if not selected_ids:
            internal_errors.append(f"Debes seleccionar al menos un fondo de {block}.")
            continue

        st.caption(f"Distribución interna de {block}. Debe sumar 100%.")
        if st.button(f"Repartir {block} a partes iguales", key=f"equal_weight_{block}"):
            base_weight = int(100 // len(selected_ids))
            remainder = int(100 - base_weight * len(selected_ids))
            for j, fund_id in enumerate(selected_ids):
                st.session_state[f"weight_{block}_{fund_id}"] = base_weight + (remainder if j == 0 else 0)
            for k in ["result", "final_weights_df", "profile"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()

        internal_weights = {}
        cols = st.columns(2)
        for i, identifier in enumerate(selected_ids):
            label = all_id_to_label.get(identifier, identifier)
            with cols[i % 2]:
                internal_weights[identifier] = st.slider(
                    label,
                    min_value=0,
                    max_value=100,
                    value=int(round(100 / len(selected_ids))),
                    step=5,
                    key=f"weight_{block}_{identifier}",
                )

        internal_sum = sum(internal_weights.values())
        st.write(f"Suma interna {block}: **{internal_sum:.0f}%**")

        if internal_sum != 100:
            internal_errors.append(f"La suma interna de {block} debe ser 100%.")
            continue

        for identifier, internal_weight in internal_weights.items():
            row = funds_block[funds_block["Identificador"] == identifier].iloc[0]
            final_weight = target_weight * internal_weight / 100
            final_rows.append(
                {
                    "Identificador": identifier,
                    "Nombre": row["Nombre"],
                    "Bloque": block,
                    "Categoría": row["Categoría"],
                    "Subcategoría": row["Subcategoría"],
                    "Peso interno": internal_weight,
                    "Peso final": final_weight,
                }
            )

    if internal_errors:
        for err in internal_errors:
            st.error(err)
        st.stop()

    final_weights_df = pd.DataFrame(final_rows)
    if final_weights_df.empty:
        st.stop()

    st.subheader("4. Cartera final")
    display_weights = final_weights_df.copy()
    display_weights["Peso final"] = display_weights["Peso final"].map(lambda x: f"{x:.2f}%")
    st.dataframe(display_weights, use_container_width=True, hide_index=True)

    final_weights = dict(zip(final_weights_df["Identificador"], final_weights_df["Peso final"]))

    if st.button("Calcular backtest", type="primary"):
        try:
            result = calculate_backtest(navs, final_weights)
            st.session_state["result"] = result
            st.session_state["final_weights_df"] = final_weights_df
            st.session_state["profile"] = profile
        except Exception as e:
            st.error(f"No se ha podido calcular el backtest: {e}")
            st.stop()

    if "result" in st.session_state:
        result = st.session_state["result"]
        final_weights_df = st.session_state["final_weights_df"]
        profile = st.session_state["profile"]

        st.divider()
        st.subheader("5. Informe de cartera")

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Rentabilidad acumulada", format_pct(result["total_return"]))
        m2.metric("Rentabilidad anualizada", format_pct(result["annual_return"]))
        m3.metric("Volatilidad anualizada", format_pct(result["annual_volatility"]))
        m4.metric("Ratio Sharpe", format_num(result["sharpe"]))
        m5.metric("Máx. drawdown", format_pct(result["max_drawdown"]))

        liquidativo_df = pd.DataFrame(
            {
                "Mes": np.arange(1, len(result["liquidativo"]) + 1),
                "Liquidativo": result["liquidativo"].values,
            }
        )
        tickvals, ticktext = make_year_axis(liquidativo_df, "Mes")

        fig = px.line(
            liquidativo_df,
            x="Mes",
            y="Liquidativo",
            title="Evolución del liquidativo base 10.000 — fechas ocultas",
            template="plotly_white",
        )
        fig.update_traces(line_width=3)
        fig.update_layout(
            height=500,
            xaxis_title="Año de simulación",
            yaxis_title="Liquidativo base 10.000",
            title_x=0.02,
            xaxis=dict(tickmode="array", tickvals=tickvals, ticktext=ticktext),
        )
        st.plotly_chart(fig, use_container_width=True)

        drawdown_df = pd.DataFrame(
            {
                "Mes": np.arange(1, len(result["drawdown"]) + 1),
                "Drawdown": result["drawdown"].values,
            }
        )
        tickvals_dd, ticktext_dd = make_year_axis(drawdown_df, "Mes")
        fig_dd = px.area(
            drawdown_df,
            x="Mes",
            y="Drawdown",
            title="Drawdown de la cartera",
            template="plotly_white",
        )
        fig_dd.update_layout(
            height=320,
            xaxis_title="Año de simulación",
            yaxis_title="Drawdown",
            yaxis_tickformat=".0%",
            title_x=0.02,
            xaxis=dict(tickmode="array", tickvals=tickvals_dd, ticktext=ticktext_dd),
        )
        st.plotly_chart(fig_dd, use_container_width=True)

        st.subheader("Comportamiento de los fondos seleccionados")
        st.caption("Métricas calculadas para cada fondo en el mismo periodo oculto del backtest. No se muestran fechas reales.")
        fund_metrics_display = prepare_fund_metrics_table(result, final_weights_df)
        if not fund_metrics_display.empty:
            st.dataframe(
                format_fund_metrics_for_display(fund_metrics_display),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("No hay métricas individuales disponibles para los fondos seleccionados.")

        report_col1, report_col2 = st.columns(2)
        with report_col1:
            one_pager = create_one_pager_excel(result, final_weights_df, profile)
            st.download_button(
                "Descargar one pager: Informe de cartera",
                data=one_pager,
                file_name="informe_de_cartera.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with report_col2:
            try:
                student_report_row = pd.Series({
                    "name": user_info.get("name", username),
                    "internal_id": user_info.get("internal_id", ""),
                    "profile": profile,
                    "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M"),
                })
                pdf_bytes = create_portfolio_pdf_report(student_report_row, final_weights_df, result)
                st.download_button(
                    "Descargar informe de cartera en PDF",
                    data=pdf_bytes,
                    file_name="informe_de_cartera.pdf",
                    mime="application/pdf",
                )
            except ModuleNotFoundError:
                st.warning("Para generar informes PDF instala ReportLab: python3 -m pip install reportlab")
            except Exception as e:
                st.warning(f"No se ha podido generar el PDF: {e}")

        col_save, col_note = st.columns([1, 3])
        with col_save:
            if st.button("Guardar entrega"):
                save_submission(username, user_info, profile, final_weights_df, result)
                st.success("Cartera guardada correctamente.")
        with col_note:
            st.caption("El profesor podrá consultar las carteras guardadas desde el panel de administración.")


# ------------------------------------------------------------
# Vista profesor
# ------------------------------------------------------------
def teacher_page(username: str, user_info: dict):
    st.markdown(
        """
        <div class="main-header">
            <h1>Panel del profesor</h1>
            <p>Consulta perfiles, resultados y la cartera completa de cada alumno.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("Administración de usuarios")
    users = load_users()

    with st.expander("Crear nuevo usuario", expanded=False):
        with st.form("create_user_form"):
            c1, c2 = st.columns(2)
            with c1:
                new_username = st.text_input("Usuario", placeholder="alumno03")
                new_name = st.text_input("Nombre visible", placeholder="Alumno 3")
                new_internal_id = st.text_input("ID interno", placeholder="A003")
            with c2:
                new_role = st.selectbox("Rol", ["alumno", "profesor"], index=0)
                new_password = st.text_input("Contraseña inicial", type="password")
                repeat_password = st.text_input("Repetir contraseña", type="password")

            submitted = st.form_submit_button("Crear usuario", type="primary")
            if submitted:
                if new_password != repeat_password:
                    st.error("Las contraseñas no coinciden.")
                else:
                    ok, msg = create_user(new_username, new_password, new_name, new_internal_id, new_role)
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    users_rows = []
    profiles = load_user_profiles()
    for user, data in users.items():
        profile_data = profiles.get(user, {})
        users_rows.append({
            "usuario": user,
            "nombre": data.get("name", user),
            "id_interno": data.get("internal_id", ""),
            "rol": data.get("role", ""),
            "activo": "Sí" if data.get("active", True) else "No",
            "perfil_test": profile_data.get("profile", ""),
            "puntuación_test": profile_data.get("score", ""),
        })

    if users_rows:
        st.dataframe(pd.DataFrame(users_rows), use_container_width=True, hide_index=True)

    with st.expander("Gestionar usuario existente", expanded=False):
        user_options = sorted(users.keys())
        selected_user_admin = st.selectbox("Selecciona usuario", user_options, key="admin_selected_user")
        if selected_user_admin:
            selected_meta = users[selected_user_admin]
            st.write(f"**Nombre:** {selected_meta.get('name', selected_user_admin)}")
            st.write(f"**Rol:** {selected_meta.get('role', '')}")
            st.write(f"**Estado:** {'Activo' if selected_meta.get('active', True) else 'Desactivado'}")

            ac1, ac2, ac3 = st.columns(3)
            with ac1:
                new_pass_admin = st.text_input("Nueva contraseña", type="password", key="admin_new_password")
                if st.button("Cambiar contraseña", key="admin_change_password"):
                    ok, msg = update_user_password(selected_user_admin, new_pass_admin)
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
            with ac2:
                if selected_meta.get("active", True):
                    if st.button("Desactivar usuario", key="admin_deactivate_user"):
                        ok, msg = set_user_active(selected_user_admin, False)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                else:
                    if st.button("Reactivar usuario", key="admin_activate_user"):
                        ok, msg = set_user_active(selected_user_admin, True)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
            with ac3:
                if st.button("Resetear test de perfil", key="admin_reset_profile"):
                    delete_user_risk_profile(selected_user_admin)
                    st.success("Test de perfil reseteado para ese usuario.")
                    st.rerun()
                confirm_delete = st.checkbox("Confirmar borrado", key="admin_confirm_delete")
                if st.button("Borrar usuario", key="admin_delete_user"):
                    if not confirm_delete:
                        st.error("Marca 'Confirmar borrado' antes de borrar el usuario.")
                    else:
                        ok, msg = delete_user_completely(selected_user_admin)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)

    st.divider()

    profiles = load_user_profiles()
    if profiles:
        st.subheader("Perfiles de riesgo guardados")
        profile_rows = []
        for user, data in profiles.items():
            user_meta = users.get(user, {})
            profile_rows.append(
                {
                    "usuario": user,
                    "nombre": user_meta.get("name", user),
                    "id_interno": user_meta.get("internal_id", ""),
                    "perfil": data.get("profile", ""),
                    "puntuación": data.get("score", ""),
                    "fecha_test": data.get("timestamp", ""),
                }
            )
        st.dataframe(pd.DataFrame(profile_rows), use_container_width=True, hide_index=True)
    else:
        st.info("Todavía no hay perfiles de riesgo guardados.")

    submissions = load_submissions()
    if submissions.empty:
        st.info("Todavía no hay carteras guardadas por alumnos.")
        return

    # Intenta cargar clasificación e histórico para enriquecer carteras antiguas y generar informes PDF.
    classification = pd.DataFrame()
    navs_for_teacher = pd.DataFrame()
    try:
        classification, navs_for_teacher, _ = load_database_from_excel(None, str(DEFAULT_EXCEL))
    except Exception:
        classification = pd.DataFrame()
        navs_for_teacher = pd.DataFrame()

    st.subheader("Resumen de entregas")
    display_cols = [
        "timestamp",
        "username",
        "internal_id",
        "name",
        "profile",
        "num_fondos",
        "total_return",
        "annual_return",
        "annual_volatility",
        "sharpe",
        "max_drawdown",
    ]
    existing_display_cols = [c for c in display_cols if c in submissions.columns]
    summary = submissions[existing_display_cols].copy()

    for col in ["total_return", "annual_return", "annual_volatility", "max_drawdown"]:
        if col in summary.columns:
            summary[col] = summary[col].map(lambda x: f"{x:.2%}" if pd.notna(x) else "")
    if "sharpe" in summary.columns:
        summary["sharpe"] = summary["sharpe"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "")

    st.dataframe(summary, use_container_width=True, hide_index=True)

    st.download_button(
        "Descargar todas las entregas en CSV",
        data=submissions.to_csv(index=False).encode("utf-8"),
        file_name="entregas_carteras_alumnos.csv",
        mime="text/csv",
    )

    st.subheader("Ranking por ratio Sharpe")
    if "sharpe" in submissions.columns:
        ranking = submissions.sort_values("sharpe", ascending=False)
        st.dataframe(ranking[existing_display_cols], use_container_width=True, hide_index=True)
    else:
        st.info("No hay columna Sharpe en las entregas guardadas.")

    st.subheader("Cartera completa por alumno")
    if "name" not in submissions.columns:
        st.info("No encuentro la columna de alumno en las entregas.")
        return

    # Selector robusto: alumno + fecha de entrega, para diferenciar varias entregas del mismo alumno.
    submissions_for_selector = submissions.reset_index().copy()
    submissions_for_selector["selector"] = submissions_for_selector.apply(
        lambda r: f"{r.get('name', '')} | {r.get('internal_id', '')} | {r.get('timestamp', '')}",
        axis=1,
    )

    selected_submission = st.selectbox(
        "Selecciona una entrega",
        submissions_for_selector["selector"].tolist(),
    )

    selected_index = submissions_for_selector.loc[
        submissions_for_selector["selector"] == selected_submission, "index"
    ].iloc[0]
    selected_row = submissions.loc[selected_index]

    detail_metrics_cols = [
        "timestamp",
        "username",
        "internal_id",
        "name",
        "profile",
        "total_return",
        "annual_return",
        "annual_volatility",
        "sharpe",
        "max_drawdown",
    ]
    detail_metrics = pd.DataFrame(
        [
            {
                col: selected_row[col]
                for col in detail_metrics_cols
                if col in selected_row.index
            }
        ]
    )
    for col in ["total_return", "annual_return", "annual_volatility", "max_drawdown"]:
        if col in detail_metrics.columns:
            detail_metrics[col] = detail_metrics[col].map(lambda x: f"{x:.2%}" if pd.notna(x) else "")
    if "sharpe" in detail_metrics.columns:
        detail_metrics["sharpe"] = detail_metrics["sharpe"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "")

    st.markdown("**Métricas de la entrega seleccionada**")
    st.dataframe(detail_metrics, use_container_width=True, hide_index=True)

    portfolio_detail = portfolio_from_submission_row(selected_row, classification)
    if portfolio_detail.empty:
        st.warning("No se ha podido reconstruir la cartera de esta entrega.")
        return

    portfolio_detail = portfolio_detail.copy()
    portfolio_detail["Peso final"] = pd.to_numeric(portfolio_detail["Peso final"], errors="coerce").fillna(0.0)
    portfolio_detail = portfolio_detail.sort_values("Peso final", ascending=False)

    st.markdown("**Fondos seleccionados y pesos**")
    portfolio_display = portfolio_detail.copy()
    portfolio_display["Peso final"] = portfolio_display["Peso final"].map(lambda x: f"{x:.2f}%")

    cols_to_show = [
        "Identificador",
        "Nombre",
        "Bloque",
        "Categoría",
        "Subcategoría",
        "Peso final",
    ]
    cols_to_show = [c for c in cols_to_show if c in portfolio_display.columns]
    st.dataframe(portfolio_display[cols_to_show], use_container_width=True, hide_index=True)

    if not navs_for_teacher.empty:
        try:
            weights_for_metrics = dict(zip(portfolio_detail["Identificador"], portfolio_detail["Peso final"]))
            result_for_metrics = calculate_backtest(navs_for_teacher, weights_for_metrics)
            fund_metrics_teacher = prepare_fund_metrics_table(result_for_metrics, portfolio_detail)
            if not fund_metrics_teacher.empty:
                st.markdown("**Comportamiento de los fondos en el periodo analizado**")
                st.caption("Rentabilidad acumulada, rentabilidad anualizada y volatilidad anualizada de cada fondo. Fechas reales omitidas.")
                st.dataframe(
                    format_fund_metrics_for_display(fund_metrics_teacher),
                    use_container_width=True,
                    hide_index=True,
                )
        except Exception as e:
            st.info(f"No se han podido calcular las métricas individuales de fondos: {e}")

    # Resumen por grandes bloques y categorías.
    st.markdown("**Resumen por bloques y categorías**")
    c1, c2 = st.columns(2)
    with c1:
        if "Bloque" in portfolio_detail.columns:
            by_block = (
                portfolio_detail.groupby("Bloque", dropna=False)["Peso final"]
                .sum()
                .reset_index()
                .sort_values("Peso final", ascending=False)
            )
            by_block["Peso final"] = by_block["Peso final"].map(lambda x: f"{x:.2f}%")
            st.dataframe(by_block, use_container_width=True, hide_index=True)
    with c2:
        if "Categoría" in portfolio_detail.columns:
            by_category = (
                portfolio_detail.groupby("Categoría", dropna=False)["Peso final"]
                .sum()
                .reset_index()
                .sort_values("Peso final", ascending=False)
            )
            by_category["Peso final"] = by_category["Peso final"].map(lambda x: f"{x:.2f}%")
            st.dataframe(by_category, use_container_width=True, hide_index=True)

    export_portfolio = portfolio_detail.copy()
    export_portfolio.insert(0, "alumno", selected_row.get("name", ""))
    export_portfolio.insert(1, "id_interno", selected_row.get("internal_id", ""))
    export_portfolio.insert(2, "fecha_entrega", selected_row.get("timestamp", ""))

    st.download_button(
        "Descargar cartera seleccionada en CSV",
        data=export_portfolio.to_csv(index=False).encode("utf-8"),
        file_name="cartera_completa_alumno.csv",
        mime="text/csv",
    )

    st.markdown("**Informe PDF completo**")
    if navs_for_teacher.empty:
        st.warning("No se ha podido cargar el histórico de liquidativos para generar el PDF completo.")
    else:
        try:
            weights_for_pdf = dict(zip(portfolio_detail["Identificador"], portfolio_detail["Peso final"]))
            result_for_pdf = calculate_backtest(navs_for_teacher, weights_for_pdf)
            pdf_bytes = create_portfolio_pdf_report(selected_row, portfolio_detail, result_for_pdf)
            safe_name = str(selected_row.get("name", "alumno")).strip().replace(" ", "_").lower() or "alumno"
            st.download_button(
                "Descargar informe completo en PDF",
                data=pdf_bytes,
                file_name=f"informe_cartera_{safe_name}.pdf",
                mime="application/pdf",
            )
        except ModuleNotFoundError:
            st.warning("Para generar informes PDF instala ReportLab: python3 -m pip install reportlab")
        except Exception as e:
            st.warning(f"No se ha podido generar el PDF de esta cartera: {e}")


# ------------------------------------------------------------
# App principal
# ------------------------------------------------------------
def main():
    if not st.session_state.authenticated:
        login_page()
        return

    global USERS
    USERS = load_users()
    username = st.session_state.username
    if username not in USERS or not USERS[username].get("active", True):
        st.session_state.authenticated = False
        st.session_state.username = None
        st.error("El usuario ya no existe o ha sido desactivado.")
        return
    user_info = USERS[username]

    with st.sidebar:
        st.title("Menú")
        st.write(f"**Usuario:** {user_info['name']}")
        st.write(f"**Rol:** {user_info['role']}")

        if user_info["role"] == "alumno":
            profile_data = get_user_risk_profile(username)
            if profile_data:
                st.write(f"**Perfil:** {profile_data.get('profile', 'n.d.')}")
                st.write(f"**Puntuación:** {profile_data.get('score', 'n.d.')}")
                with st.expander("Opciones del test"):
                    if st.button("Rehacer test de perfil"):
                        delete_user_risk_profile(username)
                        for key in list(st.session_state.keys()):
                            if key.startswith("selected_ids_") or key.startswith("weight_") or key in ["result", "final_weights_df", "profile"]:
                                del st.session_state[key]
                        st.rerun()

        if st.button("Cerrar sesión"):
            st.session_state.authenticated = False
            st.session_state.username = None
            st.rerun()

    if user_info["role"] == "profesor":
        teacher_page(username, user_info)
    else:
        if get_user_risk_profile(username) is None:
            risk_profile_test_page(username, user_info)
        else:
            student_page(username, user_info)


if __name__ == "__main__":
    main()
