from http.server import BaseHTTPRequestHandler
import json
import math
import os
import pandas as pd
import requests
from urllib.parse import parse_qs, urlparse


def safe_round(v, digits=2):
    """Convierte NaN/Inf a None (→ null en JSON) y redondea el resto."""
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f)) else round(f, digits)
    except (TypeError, ValueError):
        return None

# ── Constantes ──────────────────────────────────────────────────────────────

COLUMNAS = [
    'fecha', 'Indice_IPC', 'Indice_Salarial', 'Costo_Canasta',
    'Alimentos y bebidas no alcohólicas',
    'Bebidas alcohólicas y tabaco',
    'Prendas de vestir y calzado',
    'Vivienda, agua, electricidad, gas y otros combustibles',
    'Equipamiento y mantenimiento del hogar',
    'Salud', 'Transporte', 'Comunicaciones', 'Recreación y cultura',
    'Educación', 'Restaurantes y hoteles', 'Bienes y servicios varios',
    'Privado', 'Público', 'Informal',
]

API_INDEC = (
    "https://apis.datos.gob.ar/series/api/series/"
    "?ids=145.3_INGNACNAL_DICI_M_15,149.1_TL_INDIIOS_OCTU_0_21,"
    "150.1_CSTA_BATAL_0_D_20,146.3_IALIMENNAL_DICI_M_45,"
    "146.3_IBEBIDANAL_DICI_M_39,146.3_IPRENDANAL_DICI_M_35,"
    "146.3_IVIVIENNAL_DICI_M_52,146.3_IEQUIPANAL_DICI_M_46,"
    "146.3_ISALUDNAL_DICI_M_18,146.3_ITRANSPNAL_DICI_M_23,"
    "146.3_ICOMUNINAL_DICI_M_27,146.3_IRECREANAL_DICI_M_31,"
    "146.3_IEDUCACNAL_DICI_M_22,146.3_IRESTAUNAL_DICI_M_33,"
    "146.3_IBIENESNAL_DICI_M_36,149.1_SOR_PRIADO_OCTU_0_25,"
    "149.1_SOR_PUBICO_OCTU_0_14,149.1_SOR_PRIADO_OCTU_0_28"
    "&limit=5000&start_date=2016-12-01&format=json"
)

# Mapa: query param → columna en el DataFrame
SALARIO_MAP = {
    'privado':  'Privado',
    'publico':  'Público',
    'informal': 'Informal',
    'general':  'Indice_Salarial',
}

CLASES_VALIDAS = {'Alta', 'Media', 'Baja'}

# ── Helpers de datos ─────────────────────────────────────────────────────────

def load_ponderaciones():
    """Carga el Excel de ponderaciones desde la raíz del proyecto."""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(root, 'ponderaciones.xlsx')
    df = pd.read_excel(path)
    df.set_index('Rubro', inplace=True)
    return df


def fetch_indec():
    """Descarga los índices del INDEC y devuelve un DataFrame indexado por fecha."""
    resp = requests.get(API_INDEC, timeout=30)
    resp.raise_for_status()
    datos = resp.json()
    df = pd.DataFrame(datos['data'], columns=COLUMNAS)
    df.set_index('fecha', inplace=True)
    df.index = pd.to_datetime(df.index)
    df = df.astype(float)
    return df


def compute_isr(df_all, df_pond, clase, salario_col, fecha_inicio):
    """
    Calcula el ISR (Índice de Salario Real) para la clase y tipo salarial dados.
    Devuelve DataFrame con columnas: fecha, IPC Clase, ISAL, ISR.
    """
    # Filtrar desde la fecha base
    df_rubros = df_all.drop(
        columns=['Indice_IPC', 'Indice_Salarial', 'Costo_Canasta', 'Privado', 'Público', 'Informal']
    )
    df_base = df_rubros[df_rubros.index >= fecha_inicio].copy()

    if df_base.empty:
        raise ValueError(f"No hay datos desde {fecha_inicio}.")

    # Fecha real de inicio (puede ser el primer dato disponible >= fecha_inicio)
    first_ts = df_base.index[0]

    # IPC ponderado por clase
    pesos = df_pond[clase]
    df_isr = df_base.dot(pesos).to_frame('IPC Clase')
    df_isr['IPC Clase'] = df_isr['IPC Clase'] / df_isr.loc[first_ts, 'IPC Clase'] * 100

    # Índice salarial
    isal_series = df_all.loc[df_all.index >= first_ts, salario_col]
    df_isr['ISAL'] = isal_series / isal_series.iloc[0] * 100

    # Salario real
    df_isr['ISR'] = df_isr['ISAL'] / df_isr['IPC Clase'] * 100

    # Propagar hacia adelante gaps puntuales del INDEC y eliminar NaN restantes
    df_isr = df_isr.ffill().dropna(subset=['ISR', 'ISAL', 'IPC Clase'])

    df_isr = df_isr.reset_index()
    df_isr['fecha'] = df_isr['fecha'].dt.strftime('%Y-%m-%d')
    return df_isr


def top_rubros_por_clase(df_pond, top_n=3):
    result = {}
    for clase in CLASES_VALIDAS:
        top = df_pond[clase].sort_values(ascending=False).head(top_n)
        result[clase] = [
            {'rubro': idx, 'peso': float(v)}
            for idx, v in top.items()
        ]
    return result


# ── Handler HTTP ─────────────────────────────────────────────────────────────

class handler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):  # noqa: A002
        pass  # silenciar logs de acceso en Vercel

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)

        def p(key, default=None):
            return params.get(key, [default])[0]

        try:
            df_all  = fetch_indec()
            df_pond = load_ponderaciones()

            canasta_max = float(df_all['Costo_Canasta'].dropna().max())
            fecha_min   = df_all.index.min().strftime('%Y-%m-%d')
            fecha_max   = df_all.index.max().strftime('%Y-%m-%d')

            clase       = p('clase')
            salario_key = p('salario')
            fecha_str   = p('fecha')    # formato YYYY-MM-DD

            # ── Solo metadatos (sin parámetros de análisis) ──────────────────
            if not clase or not salario_key or not fecha_str:
                self.send_json({
                    'canasta_max': canasta_max,
                    'fecha_min':   fecha_min,
                    'fecha_max':   fecha_max,
                })
                return

            # ── Validación ───────────────────────────────────────────────────
            if clase not in CLASES_VALIDAS:
                raise ValueError(f"Clase inválida: '{clase}'. Valores posibles: {sorted(CLASES_VALIDAS)}")

            salario_col = SALARIO_MAP.get(salario_key.lower())
            if not salario_col:
                raise ValueError(f"Tipo de salario inválido: '{salario_key}'")

            # Normalizar fecha a primer día del mes
            fecha_ts = pd.Timestamp(fecha_str).replace(day=1)
            if fecha_ts < df_all.index.min():
                fecha_ts = df_all.index.min()
            if fecha_ts > df_all.index.max():
                raise ValueError("La fecha de inicio supera el último dato disponible.")

            # ── Cómputo ──────────────────────────────────────────────────────
            df_isr = compute_isr(df_all, df_pond, clase, salario_col, fecha_ts)
            top_rubros = top_rubros_por_clase(df_pond)

            self.send_json({
                'fechas':      df_isr['fecha'].tolist(),
                'isr':         [safe_round(v) for v in df_isr['ISR']],
                'ipc':         [safe_round(v) for v in df_isr['IPC Clase']],
                'isal':        [safe_round(v) for v in df_isr['ISAL']],
                'fecha_min':   fecha_min,
                'fecha_max':   fecha_max,
                'canasta_max': canasta_max,
                'ultimo_isr':  safe_round(df_isr['ISR'].iloc[-1]),
                'clase':       clase,
                'salario_key': salario_key,
                'top_rubros':  top_rubros,
            })

        except Exception as exc:
            self.send_json({'error': str(exc)}, status=500)
