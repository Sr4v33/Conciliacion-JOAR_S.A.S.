"""
Punto de entrada del Robot de Conciliación DIAN–Siigo
Rectificadora JOAR S.A.S.
"""

from __future__ import annotations

import argparse
import os
import sys
import configparser
import traceback
import shutil
from datetime import datetime

# Asegurar que el directorio del ejecutable sea el directorio de trabajo
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(BASE_DIR)
sys.path.insert(0, BASE_DIR)

from utils import logger, configurar_logging
from conciliador import conciliar_inverso
from generador_excel import generar_reporte_inverso
from enviador_correo import enviar_correo_inverso


# ---------------------------------------------------------------------------
# Lectura de configuración
# ---------------------------------------------------------------------------

def leer_configuracion(ruta_config: str) -> configparser.ConfigParser:
    if not os.path.exists(ruta_config):
        raise FileNotFoundError(
            f"Archivo de configuración no encontrado: {ruta_config}\n"
            "Asegúrese de que 'config.ini' esté en el mismo directorio que el ejecutable."
        )
    cfg = configparser.ConfigParser()
    cfg.read(ruta_config, encoding="utf-8")
    return cfg


def _cfg_get(cfg: configparser.ConfigParser, seccion: str, clave: str, por_defecto: str = "") -> str:
    try:
        return cfg.get(seccion, clave).strip()
    except (configparser.NoSectionError, configparser.NoOptionError):
        logger.warning("Parámetro no encontrado en config.ini: [%s] %s. Usando: '%s'", seccion, clave, por_defecto)
        return por_defecto


# ---------------------------------------------------------------------------
# Flujo principal del Robot
# ---------------------------------------------------------------------------

def main():
    # Inicializa el logger dual (Consola + Archivo físico en ./Logs/)
    configurar_logging()

    parser = argparse.ArgumentParser(description="RPA Conciliación DIAN–Siigo | JOAR S.A.S.")
    parser.add_argument(
        "--config",
        default=os.path.join(BASE_DIR, "config.ini"),
        help="Ruta al archivo de configuración (por defecto: config.ini)",
    )
    args = parser.parse_args()

    logger.info("=" * 65)
    logger.info("  ROBOT RPA — CONCILIACIÓN DIAN–SIIGO")
    logger.info("  Rectificadora JOAR S.A.S. | NIT 901363706-7")
    logger.info("=" * 65)

    # 1. Cargar e interpretar el archivo de configuración externo
    try:
        cfg = leer_configuracion(args.config)
        logger.info("Configuración cargada con éxito desde: %s", args.config)
    except FileNotFoundError as e:
        logger.critical(str(e))
        _pausar_y_salir(1)

    # Parámetros de Rutas y Archivos
    carpeta_entrada   = _cfg_get(cfg, "archivos", "carpeta_entrada", "./input")
    carpeta_salida    = _cfg_get(cfg, "archivos", "carpeta_salida",  "./output")
    carpeta_historico = _cfg_get(cfg, "archivos", "carpeta_historico", "./Historico_Procesados")
    nombre_dian       = _cfg_get(cfg, "archivos", "archivo_dian",    "dian.xlsx")
    nombre_siigo      = _cfg_get(cfg, "archivos", "archivo_siigo",   "siigo.xlsx")

    ruta_dian   = os.path.normpath(os.path.join(BASE_DIR, carpeta_entrada, nombre_dian))
    ruta_siigo  = os.path.normpath(os.path.join(BASE_DIR, carpeta_entrada, nombre_siigo))
    ruta_output = os.path.normpath(os.path.join(BASE_DIR, carpeta_salida))
    ruta_hist   = os.path.normpath(os.path.join(BASE_DIR, carpeta_historico))

    # Asegurar la existencia de los directorios de trabajo para evitar fallos de escritura
    os.makedirs(os.path.dirname(ruta_dian), exist_ok=True)
    os.makedirs(ruta_output, exist_ok=True)
    os.makedirs(ruta_hist, exist_ok=True)

    # Parámetros internos de la lógica de Conciliación
    tipos_doc_raw = _cfg_get(
        cfg, "conciliacion", "tipos_documento_dian",
        "Factura electrónica,Documento equivalente POS,Nota de crédito electrónica",
    )
    tipos_doc = [t.strip() for t in tipos_doc_raw.split(",") if t.strip()]
    grupo_dian = _cfg_get(cfg, "conciliacion", "grupo_dian", "Recibido")

    # Parámetros de Servidor de Correos
    smtp_server  = _cfg_get(cfg, "correo", "smtp_server",      "smtp.gmail.com")
    smtp_port    = int(_cfg_get(cfg, "correo", "smtp_port",    "587"))
    remitente    = _cfg_get(cfg, "correo", "remitente",         "")
    password     = _cfg_get(cfg, "correo", "password_remitente", "")
    destinatario = _cfg_get(cfg, "correo", "destinatario",      "")
    enviar       = _cfg_get(cfg, "correo", "enviar_correo",     "false").lower() == "true"

    # 2. Ejecución del proceso de cruce de datos
    logger.info("-" * 65)
    logger.info("PASO 1/4 — Iniciando cruce y conciliación de documentos...")
    try:
        resultado, ruta_dian_marcado, ruta_siigo_marcado = conciliar_inverso(
            ruta_dian=ruta_dian,
            ruta_siigo=ruta_siigo,
            tipos_doc=tipos_doc,
            grupo=grupo_dian,
        )
    except FileNotFoundError as e:
        logger.critical(str(e))
        logger.critical("Por favor, deposite los archivos fuente en '%s' y reintente.", carpeta_entrada)
        _pausar_y_salir(1)
    except Exception as e:
        logger.critical("Error crítico inesperado en el motor de conciliación: %s", e)
        logger.debug(traceback.format_exc())
        _pausar_y_salir(1)

    # 3. Construcción del Reporte resumido
    logger.info("-" * 65)
    logger.info("PASO 2/4 — Generando el reporte de auditoría en Excel...")
    try:
        ruta_excel = generar_reporte_inverso(resultado, ruta_output)
    except Exception as e:
        logger.error("No se pudo construir el reporte Excel estructurado: %s", e)
        logger.debug(traceback.format_exc())
        ruta_excel = None

    # 4. Gestión de Históricos y Rotación Temporal (Auto-limpieza de carpetas)
    logger.info("-" * 65)
    logger.info("PASO 3/4 — Ejecutando archivado histórico y rotación de fuentes...")
    try:
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_hist_actual = os.path.join(ruta_hist, timestamp_str)
        os.makedirs(ruta_hist_actual, exist_ok=True)

        # Respaldo inmutable de los archivos cargados originalmente en esta ejecución
        if os.path.exists(ruta_dian):
            shutil.move(ruta_dian, os.path.join(ruta_hist_actual, nombre_dian))
        if os.path.exists(ruta_siigo):
            shutil.move(ruta_siigo, os.path.join(ruta_hist_actual, nombre_siigo))

        # Traslado seguro de los archivos marcados hacia la carpeta final (archivos_salida) con sello de tiempo
        if ruta_dian_marcado and os.path.exists(ruta_dian_marcado):
            nuevo_nombre_dian = f"dian_marcado_{timestamp_str}.xlsx"
            nueva_ruta_dian = os.path.join(ruta_output, nuevo_nombre_dian)
            shutil.move(ruta_dian_marcado, nueva_ruta_dian)
            ruta_dian_marcado = nueva_ruta_dian

        if ruta_siigo_marcado and os.path.exists(ruta_siigo_marcado):
            nuevo_nombre_siigo = f"siigo_marcado_{timestamp_str}.xlsx"
            nueva_ruta_siigo = os.path.join(ruta_output, nuevo_nombre_siigo)
            shutil.move(ruta_siigo_marcado, nueva_ruta_siigo)
            ruta_siigo_marcado = nueva_ruta_siigo

        logger.info("Archivos fuente respaldados correctamente en: %s", ruta_hist_actual)

    except Exception as e:
        logger.error("Ocurrió un contratiempo durante la transferencia histórica: %s", e)
        logger.debug(traceback.format_exc())

    # 5. Gestión de notificaciones
    logger.info("-" * 65)
    logger.info("PASO 4/4 — Gestionando la notificación automática por correo...")
    if not enviar:
        logger.info("Envío de correos omitido (enviar_correo = false en config.ini).")
    elif not remitente or not destinatario:
        logger.warning("Faltan credenciales o direcciones de destino obligatorias en el config.ini.")
    elif ruta_excel is None:
        logger.warning("Envío cancelado debido a la ausencia del reporte Excel.")
    else:
        enviado = enviar_correo_inverso(
            resultado=resultado,
            ruta_adjunto=ruta_excel,
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            remitente=remitente,
            password=password,
            destinatario=destinatario,
        )
        if not enviado:
            logger.warning("La entrega falló. Verifique la seguridad de la cuenta emisor y los parámetros SMTP.")

    # 6. Resumen en Consola
    logger.info("=" * 65)
    logger.info("CONCILIACIÓN FINALIZADA — %s", resultado.periodo_etiqueta)
    logger.info("  Facturas DIAN      : %d", resultado.total_dian)
    logger.info("  Facturas Siigo     : %d", resultado.total_siigo)
    logger.info("  Conciliadas ✔      : %d", resultado.total_conciliadas)
    logger.info("  Faltantes en DIAN ✘: %d", resultado.total_faltantes)
    if resultado.valor_faltantes > 0:
        logger.info(
            "  Valor faltantes    : $ {:,.0f}".format(resultado.valor_faltantes).replace(",", ".")
        )
    if ruta_excel:
        logger.info("  Reporte unificado  : %s", ruta_excel)
    if ruta_dian_marcado:
        logger.info("  DIAN Marcado       : %s", ruta_dian_marcado)
    if ruta_siigo_marcado:
        logger.info("  Siigo Marcado      : %s", ruta_siigo_marcado)
    logger.info("=" * 65)

    _pausar_y_salir(0)


def _pausar_y_salir(codigo: int):
    """Mantiene abierta la interfaz de consola si el script corre desde un binario congelado."""
    if getattr(sys, "frozen", False):
        input("\nPresione ENTER para finalizar y cerrar la ventana...")
    sys.exit(codigo)


if __name__ == "__main__":
    main()