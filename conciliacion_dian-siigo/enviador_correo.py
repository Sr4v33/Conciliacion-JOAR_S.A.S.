"""
Composición y envío del correo SMTP de notificación (conciliación inversa)
Rectificadora JOAR S.A.S.

Notifica las facturas que están en DIAN pero NO tienen registro en Siigo.
"""

from __future__ import annotations

import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from conciliador import ResultadoConciliacionInversa, RegistroFaltanteDIAN
from utils import logger, formatear_moneda


# ---------------------------------------------------------------------------
# Colores para la tabla HTML del correo  (idénticos al original)
# ---------------------------------------------------------------------------

COLOR_PRIORIDAD = {
    "Media":   "#FFF9C4",
    "Alta":    "#FFE0B2",
    "Crítica": "#FFCDD2",
}

BADGE_COLOR = {
    "Media":   "#F57F17",
    "Alta":    "#E65100",
    "Crítica": "#B71C1C",
}


def enviar_correo_inverso(
    resultado: ResultadoConciliacionInversa,
    ruta_adjunto: str,
    smtp_server: str,
    smtp_port: int,
    remitente: str,
    password: str,
    destinatario: str,
) -> bool:
    """
    Construye y envía el correo de notificación para la conciliación inversa.
    Retorna True si el envío fue exitoso, False en caso contrario.
    """
    asunto = (
        f"⚠️ Conciliación Inversa DIAN–Siigo {resultado.periodo_etiqueta} "
        f"— {resultado.total_faltantes} factura(s) sin causar en Siigo"
    )

    cuerpo_html = _construir_html(resultado, os.path.basename(ruta_adjunto))

    msg            = MIMEMultipart("mixed")
    msg["Subject"] = asunto
    msg["From"]    = remitente
    msg["To"]      = destinatario

    msg.attach(MIMEText(cuerpo_html, "html", "utf-8"))

    if os.path.exists(ruta_adjunto):
        with open(ruta_adjunto, "rb") as f:
            adjunto = MIMEBase("application", "octet-stream")
            adjunto.set_payload(f.read())
        encoders.encode_base64(adjunto)
        adjunto.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(ruta_adjunto)}",
        )
        msg.attach(adjunto)
    else:
        logger.warning("Archivo adjunto no encontrado: %s", ruta_adjunto)

    try:
        logger.info("Conectando a SMTP %s:%d ...", smtp_server, smtp_port)
        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as servidor:
            servidor.ehlo()
            servidor.starttls()
            servidor.ehlo()
            servidor.login(remitente, password)
            servidor.sendmail(remitente, [destinatario], msg.as_string())
        logger.info("Correo enviado exitosamente a: %s", destinatario)
        return True
    except smtplib.SMTPAuthenticationError:
        logger.error("Error de autenticación SMTP. Verifique las credenciales en config.ini.")
    except smtplib.SMTPException as e:
        logger.error("Error SMTP al enviar correo: %s", e)
    except Exception as e:
        logger.error("Error inesperado al enviar correo: %s", e)
    return False


# ---------------------------------------------------------------------------
# Construcción del cuerpo HTML
# ---------------------------------------------------------------------------

def _construir_html(resultado: ResultadoConciliacionInversa, nombre_archivo: str) -> str:
    tabla_faltantes = _tabla_faltantes_html(resultado.faltantes)

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <style>
    body {{ font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #333; margin: 0; padding: 0; }}
    .container {{ max-width: 900px; margin: 0 auto; padding: 24px; }}
    .header {{ background: #1F3864; color: white; padding: 20px 24px; border-radius: 6px 6px 0 0; }}
    .header h1 {{ margin: 0; font-size: 20px; }}
    .header p  {{ margin: 4px 0 0; font-size: 12px; opacity: 0.8; }}
    .section {{ background: #f9f9f9; border: 1px solid #ddd; padding: 16px 24px; margin-top: 12px; border-radius: 4px; }}
    .section h2 {{ margin: 0 0 10px; font-size: 15px; color: #1F3864; border-bottom: 2px solid #2E75B6; padding-bottom: 4px; }}
    .metrics {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 8px; }}
    .metric-box {{ flex: 1; min-width: 160px; padding: 12px; border-radius: 4px; text-align: center; }}
    .metric-box.verde  {{ background: #E8F5E9; border: 1px solid #81C784; }}
    .metric-box.rojo   {{ background: #FFEBEE; border: 1px solid #E57373; }}
    .metric-box.azul   {{ background: #E3F2FD; border: 1px solid #64B5F6; }}
    .metric-box .valor {{ font-size: 26px; font-weight: bold; }}
    .metric-box .label {{ font-size: 11px; color: #555; margin-top: 4px; }}
    .metric-box.verde .valor  {{ color: #2E7D32; }}
    .metric-box.rojo  .valor  {{ color: #C62828; }}
    .metric-box.azul  .valor  {{ color: #1565C0; }}
    .alert {{ background: #FFF3E0; border-left: 4px solid #F57C00; padding: 12px 16px; border-radius: 4px; margin-top: 8px; }}
    .alert strong {{ color: #E65100; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 8px; }}
    th {{ background: #1F3864; color: white; padding: 8px 10px; text-align: left; }}
    td {{ padding: 6px 10px; border-bottom: 1px solid #e0e0e0; }}
    .badge {{ display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: bold; color: white; }}
    .footer {{ margin-top: 20px; font-size: 11px; color: #999; text-align: center; border-top: 1px solid #eee; padding-top: 12px; }}
    .adjunto {{ background: #E8EAF6; border: 1px solid #9FA8DA; padding: 10px 16px; border-radius: 4px; margin-top: 8px; font-size: 13px; }}
    .adjunto span {{ color: #1A237E; font-weight: bold; }}
    .cufe {{ font-family: monospace; font-size: 10px; color: #555; word-break: break-all; }}
  </style>
</head>
<body>
<div class="container">

  <div class="header">
    <h1>📊 Conciliación Inversa DIAN–Siigo — {resultado.periodo_etiqueta}</h1>
    <p>Rectificadora JOAR S.A.S. | NIT 901363706-7 | {resultado.timestamp_ejecucion}</p>
  </div>

  <div class="section">
    <p>Estimado equipo administrativo,</p>
    <p>El proceso automatizado de conciliación inversa DIAN–Siigo para el período
    <strong>{resultado.periodo_etiqueta}</strong>
    ({resultado.periodo_inicio} – {resultado.periodo_fin}) ha finalizado exitosamente.
    A continuación encontrará el resumen de resultados:</p>
  </div>

  <!-- Resumen métricas -->
  <div class="section">
    <h2>📋 Resumen</h2>
    <div class="metrics">
      <div class="metric-box azul">
        <div class="valor">{resultado.total_dian}</div>
        <div class="label">Facturas reportadas en DIAN</div>
      </div>
      <div class="metric-box azul">
        <div class="valor">{resultado.total_siigo}</div>
        <div class="label">Facturas causadas en Siigo</div>
      </div>
      <div class="metric-box verde">
        <div class="valor">{resultado.total_conciliadas}</div>
        <div class="label">Facturas conciliadas ✔</div>
      </div>
      <div class="metric-box rojo">
        <div class="valor">{resultado.total_faltantes}</div>
        <div class="label">Faltantes en Siigo ✘</div>
      </div>
    </div>
  </div>

  <!-- Acción requerida -->
  <div class="section">
    <h2>⚠️ Acción requerida</h2>
    {_bloque_accion(resultado)}
  </div>

  <!-- Detalle de faltantes -->
  <div class="section">
    <h2>📄 Detalle de facturas validadas en DIAN sin registro en Siigo</h2>
    {tabla_faltantes}
  </div>

  <!-- Archivo adjunto -->
  <div class="section">
    <h2>📎 Archivo adjunto</h2>
    <div class="adjunto">
      Se adjunta el archivo Excel con el reporte completo: <span>{nombre_archivo}</span><br>
      Contiene las hojas <em>Resumen</em> y <em>Faltantes</em> con toda la información detallada.
    </div>
  </div>

  <div class="footer">
    Robot RPA — Conciliación Inversa DIAN-Siigo &nbsp;|&nbsp; Rectificadora JOAR S.A.S. &nbsp;|&nbsp;
    Ejecutado el {resultado.timestamp_ejecucion}
  </div>

</div>
</body>
</html>"""


def _bloque_accion(resultado: ResultadoConciliacionInversa) -> str:
    if resultado.total_faltantes == 0:
        return """<p style="color:#2E7D32; font-weight:bold;">
            ✔  No se requiere ninguna acción. Todas las facturas validadas en DIAN
            tienen registro correspondiente en Siigo.
        </p>"""
    return f"""<div class="alert">
        <strong>Se identificaron {resultado.total_faltantes} factura(s)</strong> validadas en DIAN
        que <strong>no tienen registro de causación en Siigo</strong>, con un valor total acumulado de
        <strong>{formatear_moneda(resultado.valor_faltantes)}</strong>.<br><br>
        Por favor revise el detalle a continuación y proceda a causar estas facturas en Siigo
        a la mayor brevedad posible para evitar inconsistencias en el cierre contable.
    </div>"""


def _tabla_faltantes_html(faltantes: list[RegistroFaltanteDIAN]) -> str:
    if not faltantes:
        return "<p style='color:#2E7D32; font-weight:bold;'>✔ No hay facturas faltantes en Siigo para este período.</p>"

    filas = ""
    for reg in faltantes:
        bg          = COLOR_PRIORIDAD.get(reg.prioridad, "#FFFFFF")
        # CUFE truncado para legibilidad en el correo (primeros 20 caracteres)
        cufe_corto  = reg.cufe[:20] + "…" if len(reg.cufe) > 20 else reg.cufe
        filas += f"""<tr style="background:{bg}">
            <td>{reg.tipo_documento}</td>
            <td class="cufe" title="{reg.cufe}">{cufe_corto}</td>
            <td>{reg.folio}</td>
            <td>{reg.prefijo}</td>
            <td>{reg.fecha_recepcion}</td>
            <td>{reg.nit_emisor}</td>
            <td>{reg.nombre_emisor}</td>
            <td style="text-align:right">{formatear_moneda(reg.iva)}</td>
            <td style="text-align:right">{formatear_moneda(reg.total)}</td>
            <td>{reg.estado}</td>
            <td>{reg.grupo}</td>
        </tr>"""

    return f"""<table>
      <thead>
        <tr>
          <th>Tipo Documento</th>
          <th>CUFE/CUDE</th>
          <th>Folio</th>
          <th>Prefijo</th>
          <th>Fecha Recepción</th>
          <th>NIT Emisor</th>
          <th>Nombre Emisor</th>
          <th>IVA</th>
          <th>Total</th>
          <th>Estado</th>
          <th>Grupo</th>
        </tr>
      </thead>
      <tbody>
        {filas}
      </tbody>
    </table>"""