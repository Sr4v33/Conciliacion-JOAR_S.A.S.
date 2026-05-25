"""
utils.py — Funciones auxiliares para el RPA de Conciliación DIAN–Siigo
Rectificadora JOAR S.A.S.
"""

import logging
import sys
import os
from datetime import datetime, date
import calendar

# ---------------------------------------------------------------------------
# Logging con Registro Histórico Físico
# ---------------------------------------------------------------------------

def configurar_logging() -> logging.Logger:
    """Configura y retorna un logger con salida dual: consola y archivo físico."""
    logger = logging.getLogger("rpa_conciliacion")
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:
        formatter = logging.Formatter(
            "[%(asctime)s] %(levelname)-8s %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )

        # 1. Handler para Consola
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

        # 2. Handler para Archivo (Logs históricos)
        try:
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
                
            logs_dir = os.path.join(base_path, "Logs")
            os.makedirs(logs_dir, exist_ok=True)
            
            file_handler = logging.FileHandler(
                os.path.join(logs_dir, "rpa_conciliacion.log"), 
                mode="a", 
                encoding="utf-8"
            )
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        except Exception as e:
            print(f"Advertencia: No se pudo crear el archivo físico de logs: {e}")

    return logger

logger = configurar_logging()


# ---------------------------------------------------------------------------
# Normalización de folios
# ---------------------------------------------------------------------------

def normalizar_folio(valor) -> str:
    if valor is None:
        return ""
    try:
        float_val = float(valor)
        if float_val == int(float_val):
            return str(int(float_val)).strip()
        return str(float_val).strip()
    except (ValueError, TypeError):
        return str(valor).strip()

def extraer_folio_siigo(factura_proveedor) -> str:
    if not factura_proveedor or str(factura_proveedor).lower() == "nan":
        return ""
    fp_str = str(factura_proveedor).strip()
    if "-" in fp_str:
        partes = fp_str.split("-")
        return partes[-1].strip()
    return fp_str

def a_fecha(valor) -> date | None:
    if not valor or str(valor).lower() in ("nan", "none"):
        return None
    for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(str(valor).strip(), formato).date()
        except ValueError:
            continue
    return None

def calcular_dias_antiguedad(fecha_str) -> int:
    fecha_doc = a_fecha(fecha_str)
    if not fecha_doc:
        return 0
    return (date.today() - fecha_doc).days

def prioridad_por_antiguedad(dias: int) -> str:
    if dias <= 15:
        return "Media"
    elif dias <= 30:
        return "Alta"
    else:
        return "Crítica"

def inferir_periodo(fechas_siigo: list) -> tuple[str, str, str]:
    fechas = [a_fecha(f) for f in fechas_siigo if a_fecha(f) is not None]
    if not fechas:
        hoy = date.today()
        año, mes = hoy.year, hoy.month
    else:
        from collections import Counter
        conteo = Counter((f.year, f.month) for f in fechas)
        año, mes = conteo.most_common(1)[0][0]

    nombres_meses = [
        "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    ]
    etiqueta = f"{nombres_meses[mes]} {año}"
    ultimo_dia = calendar.monthrange(año, mes)[1]
    fecha_inicio = f"01/{mes:02d}/{año}"
    fecha_fin = f"{ultimo_dia}/{mes:02d}/{año}"
    
    return etiqueta, fecha_inicio, fecha_fin

def formatear_moneda(valor: float) -> str:
    return "$ {:,.0f}".format(valor).replace(",", ".")