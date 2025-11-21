# reportes_aulas.py
import argparse
from pathlib import Path
from datetime import datetime, date
import pandas as pd
import re

BRAND = {
    "primary": "#003366",
    "primary_dark": "#0c2c55",
    "accent": "#F5B301",
    "ok_bg": "#e2f5e9",
    "ok_fg": "#1f6d3a",
    "bad_bg": "#fde2e2",
    "bad_fg": "#8a1f2b",
    "muted_bg": "#f2f4f8",
    "muted_fg": "#445566",
    "table_border": "#dee4f2",
    "panel_bg": "#f8f9fc",
    "zebra": ("#ffffff", "#f9fbff"),
}

FECHA_ETQ = date.today().strftime("%Y-%m-%d")

SUBJECT_DOCENTE   = f"Informe final – M2 (Alistamiento + Ejecución) – {{DOCENTE_LBL}} – {FECHA_ETQ}"
SUBJECT_PROGRAMA  = f"Informe final – M2 (Alistamiento + Ejecución) – {{PROGRAMA}} – {FECHA_ETQ}"
SUBJECT_GLOBAL    = f"Informe Global – M2 (Alistamiento + Ejecución) – Rectoría Centro Sur – {FECHA_ETQ}"

BOOKING_URL = ("https://outlook.office.com/bookwithme/user/"
               "56cf01a4fb97453195dc6e912f82b2a5@uniminuto.edu/meetingtype/OLZ8ynZ2zkCBBRiqMRB-aQ2"
               "?bookingcode=3df0e75a-f6ba-4539-979f-a8276c1d0fc5&anonymous&ismsaljsauthenabled&ep=mlink")

DEFAULT_DOCENTE_ATTACH = ("Circular No.12_VAC_Lineamientos para el uso y apropiacion de recursos "
                          "educativos de apoyo y campus virtual.pdf")

PDFKIT_AVAILABLE = False
PDFKIT_CONFIG = None
try:
    import pdfkit
    PDFKIT_CONFIG = pdfkit.configuration(
        wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
    )
    PDFKIT_AVAILABLE = True
except Exception:
    PDFKIT_AVAILABLE = False

PDF_OPTIONS = {"encoding": "UTF-8", "quiet": "", "enable-local-file-access": ""}


def wrap_for_pdf(html_inner: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="es"><head>
<meta charset="utf-8"/>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<style>
  html,body{{-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
  body,table,td,th{{font-family:"Segoe UI","Noto Sans","DejaVu Sans",Arial,sans-serif;font-size:14px;color:#222;line-height:1.3;}}
  thead th,.thead-th th{{background:{BRAND["primary_dark"]}!important;color:#fff!important;}}
  table{{table-layout:fixed;border-collapse:collapse;}}
  td,th{{font-variant-numeric:tabular-nums;word-break:break-word;white-space:normal;}}
  .badge-pill{{border-radius:999px;padding:4px 10px;font-weight:700;display:inline-block;font-size:12px;}}
  .rev-chip{{display:inline-block;font-size:12px;border-radius:999px;padding:3px 10px;border:1px solid #d7dde9;line-height:1.2;white-space:normal;word-break:break-word;overflow-wrap:anywhere;}}
  .rev-ok{{background:#e2f5e9;color:#1f6d3a;border-color:#cfead7;}}
  .rev-muted{{background:{BRAND["muted_bg"]};color:{BRAND["muted_fg"]};}}
  .rev-dot{{display:inline-block;width:8px;height:8px;border-radius:999px;margin-right:6px;background:#6c757d;vertical-align:middle;}}
  .rev-dot-ok{{background:#1f6d3a;}}
</style></head><body>
{html_inner}
</body></html>"""


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_email(x):
    return bool(EMAIL_RE.match(str(x)))


def to_int_or_str(x):
    try:
        return int(float(x))
    except Exception:
        return x


def parse_emails(raw: str):
    if not raw:
        return []
    sep = ';' if ';' in raw else ','
    return [e.strip() for e in raw.split(sep) if e.strip()]


def resolve_existing_paths(paths):
    out = []
    for p in (paths or []):
        try:
            q = Path(p).expanduser().resolve()
        except Exception:
            q = Path(p)
        if q.exists():
            out.append(str(q))
        else:
            print(f"⚠️ Archivo para adjuntar no existe: {p}")
    return out


def nrc_to_str(val) -> str:
    s = str(val).strip()
    try:
        if s.replace('.', '', 1).isdigit():
            return str(int(float(s)))
    except Exception:
        pass
    return s


# --- Desempeño cualitativo basado en CALIFICACION FINAL (0-100) ---

def final_qual(score):
    try:
        x = float(score)
    except Exception:
        x = 0.0
    if x >= 91:
        return ("Desempeño excelente", "EXCELENTE", "#14532d", "#dcfce7")
    elif x >= 80:
        return ("Desempeño bueno", "BUENO", "#1d4ed8", "#dbeafe")
    elif x >= 70:
        return ("Desempeño aceptable", "ACEPTABLE", "#92400e", "#ffedd5")
    else:
        return ("Desempeño insatisfactorio", "INSATISFACTORIO", "#7f1d1d", "#fee2e2")


def leyenda_html_final():
    return ("Desempeño excelente (91–100) · bueno (80–90) · "
            "aceptable (70–79) · insatisfactorio (0–69)")


def observacion_badge(texto: str) -> str:
    """
    Badge para la columna Revisión (Outlook-friendly).
    """
    t = (texto or "").strip().upper().replace("MUESTRO", "MUESTREO")
    if not t:
        return "—"
    if t.startswith("REV"):
        t = "SELECCIONADA PARA EL MUESTREO"
    is_no = "NO SELECCIONADA" in t
    is_yes = ("SELECCIONADA" in t) and ("NO" not in t)

    if is_yes:
        return (f"<span class='rev-chip rev-ok'>"
                f"<span class='rev-dot rev-dot-ok'></span>Muestreo: seleccionada</span>")
    if is_no:
        return (f"<span class='rev-chip rev-muted'>"
                f"<span class='rev-dot'></span>Muestreo: no seleccionada</span>")
    return (f"<span class='rev-chip rev-muted'>"
            f"<span class='rev-dot'></span>{texto}</span>")


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()

    for c in ["CALIFICACION", "CALIFICACION 2", "CALIFICACION FINAL"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    if "NRC" in df.columns:
        df["NRC"] = df["NRC"].apply(nrc_to_str)

    if "ID DOCENTE" in df.columns and "NRC" in df.columns:
        sort_col = "CALIFICACION FINAL" if "CALIFICACION FINAL" in df.columns else "CALIFICACION"
        df = df.sort_values(sort_col, ascending=False) \
               .drop_duplicates(subset=["ID DOCENTE", "NRC"], keep="first")
    return df


def log_envio(logfile: Path, tipo: str, destinatarios: str, asunto: str, adjuntos: list):
    row = {
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "tipo": tipo,
        "para": destinatarios,
        "asunto": asunto,
        "adjuntos": ";".join(adjuntos or [])
    }
    pd.DataFrame([row]).to_csv(
        logfile, mode="a", index=False,
        header=not logfile.exists(), encoding="utf-8"
    )


def footer_block():
    return f"<div style='margin-top:14px;font-size:12px;color:#666;text-align:right;'>Generado el {FECHA_ETQ} – Rectoría Centro Sur</div>"


def email_shell(title_html, body_html):
    header = (
        "<div style='font-size:12px;color:#445;line-height:1.6;margin:0 0 10px 0;'>"
        "<em>Este informe corresponde al "
        "<strong>seguimiento final de sus aulas (Momento 2)</strong>, "
        "integrando la <strong>Fase de Alistamiento (50%)</strong> y la "
        "<strong>Fase de Ejecución (50%)</strong>. "
        "La calificación final se interpreta cualitativamente "
        "en niveles de desempeño (excelente, bueno, aceptable e insatisfactorio).</em>"
        "</div>"
    )
    return f"""<div><span class="preheader">Informe final de seguimiento – Momento 2 (Alistamiento + Ejecución).</span></div>
<table style="background:#f2f4f8;" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr><td align="center" style="padding:28px 12px;">
    <table width="720" style="max-width:720px;background:#ffffff;border-radius:12px;box-shadow:0 4px 12px rgba(0,0,0,.08);">
      <tr><td style="background:{BRAND['accent']};height:8px;border-top-left-radius:12px;border-top-right-radius:12px;font-size:0;line-height:0;">&nbsp;</td></tr>
      <tr><td style="background:{BRAND['primary']};color:#fff;padding:18px 24px;border-bottom:1px solid #002b55;font-family:Segoe UI,Arial;">
        <div style="font-size:22px;font-weight:700;">{title_html}</div></td></tr>
      <tr><td style="padding:18px 24px;font-family:Segoe UI,Arial;color:#222;font-size:15px;line-height:1.6;">
        {header}{body_html}{footer_block()}</td></tr>
    </table></td></tr></table>"""


# ---------- DOCENTES ----------

def tabla_docente(rows):
    body_rows = []
    for r in rows:
        fase1 = r.get("CALIFICACION", 0)
        fase2 = r.get("CALIFICACION 2", 0)
        final = r.get("CALIFICACION FINAL", 0)
        desc, short, fg, bg = final_qual(final)
        puntajes_html = (
            f"Alistamiento: {to_int_or_str(fase1)}<br>"
            f"Ejecución: {to_int_or_str(fase2)}<br>"
            f"<strong>Final: {to_int_or_str(final)}</strong>"
        )
        rev = observacion_badge(r.get("OBSERVACION", ""))

        body_rows.append(f"""
        <tr style="background:{BRAND['zebra'][0]};">
          <td style="padding:10px;width:12%;text-align:center;">{r['NRC']}</td>
          <td style="padding:10px;width:40%;text-align:left;word-break:break-word;white-space:normal;overflow-wrap:anywhere;">{r.get('ASIGNATURA','')}</td>
          <td style="padding:10px;width:20%;text-align:left;word-break:break-word;white-space:normal;overflow-wrap:anywhere;">{r.get('PROGRAMA','')}</td>
          <td style="padding:10px;width:14%;text-align:left;font-size:12px;line-height:1.4;">{puntajes_html}</td>
          <td style="padding:10px;width:14%;text-align:center;">
            <span style="display:inline-block;padding:4px 10px;border-radius:999px;background:{fg};color:#fff;font-size:12px;font-weight:600;" class="badge-pill">
              {desc}
            </span>
          </td>
          <td style="padding:10px;width:14%;text-align:left;white-space:normal;word-break:break-word;overflow-wrap:anywhere;line-height:1.35;">{rev}</td>
        </tr>""")

    inner = f"""
    <table style="background:#fbfbfe;border:1px solid #e3e8f1;border-left:5px solid {BRAND['accent']};border-radius:8px;" width="100%">
      <tr><td style="padding:14px 18px;color:{BRAND['primary']};font-size:15px;font-weight:600;font-family:Segoe UI,Arial;">Resumen final de aulas revisadas (Alistamiento + Ejecución)</td></tr>
      <tr><td style="padding:0 18px 16px 18px;">
        <table width="100%" style="border-collapse:collapse;table-layout:fixed;font-family:Segoe UI,Arial;font-size:14px;border:1px solid {BRAND['table_border']};">
          <thead class="thead-th">
            <tr style="background:{BRAND['primary']};color:#fff;">
              <th style="padding:10px;text-align:center;width:12%;color:#fff!important;">NRC</th>
              <th style="padding:10px;text-align:left;width:40%;color:#fff!important;">Asignatura</th>
              <th style="padding:10px;text-align:left;width:20%;color:#fff!important;">Programa</th>
              <th style="padding:10px;text-align:left;width:14%;color:#fff!important;">Puntajes (Fase 1 y 2)</th>
              <th style="padding:10px;text-align:center;width:14%;color:#fff!important;">Desempeño final</th>
              <th style="padding:10px;text-align:center;width:14%;color:#fff!important;">Revisión</th>
            </tr>
          </thead>
          <tbody>{''.join(body_rows)}</tbody>
        </table>
      </td></tr>
      <tr><td style="padding:0 18px 14px 18px;color:#555;font-size:12px;">{leyenda_html_final()}</td></tr>
    </table>"""
    return inner


def saludo_docente(nombre, docente_id):
    nombre_lbl = nombre if (nombre and str(nombre).strip()) else "Docente"
    id_lbl = f" (ID: {to_int_or_str(docente_id)})" if docente_id not in (None, "") else ""
    return f"<p><strong>Cordial saludo, {nombre_lbl}{id_lbl},</strong></p>"


def bloque_mensaje_final_docente(rows):
    finals = [float(r.get("CALIFICACION FINAL", 0)) for r in rows]
    if not finals:
        return ""
    prom_final = sum(finals) / len(finals)
    desc, short, fg, bg = final_qual(prom_final)

    return (
        f"<p>El promedio final de sus aulas (Alistamiento + Ejecución) es "
        f"<strong>{round(prom_final,1)}</strong>, con un "
        f"<span style='background:{bg};color:{fg};padding:3px 9px;border-radius:999px;font-weight:600;font-size:12px;'>{desc}</span>.</p>"
        "<p>Recuerde que esta nota integra los dos momentos de seguimiento sobre la calidad del aula. "
        "Si desea revisar en detalle algún caso puntual o recibir retroalimentación personalizada, "
        "puede agendar un espacio conmigo.</p>"
    )


def boton_agendar():
    return (f"<div style='text-align:center;margin-top:14px;'>"
            f"<a href='{BOOKING_URL}' target='_blank' rel='noopener' "
            f"style='background:#FFD000;color:{BRAND['primary']};text-decoration:none;padding:16px 28px;"
            f"border-radius:56px;font-family:Segoe UI,Arial;font-size:16px;font-weight:800;display:inline-block;"
            f"border:3px solid {BRAND['accent']};box-shadow:0 4px 10px rgba(0,0,0,0.25);'>"
            f"📅 Agendar llamada / videollamada</a></div>")


def html_docente(nombre, docente_id, rows):
    """
    Informe por docente.
    Incluye:
      - Saludo y contexto
      - KPIs de sus aulas
      - Barra horizontal apilada (distribución excelente/bueno/aceptable/insatisfactorio)
      - Tabla por NRC
      - Mensaje final y botón para agendar
    Estilo unificado con el informe global y el informe de programas.
    """
    # ---- KPIs y distribución para el docente ----
    finals = [float(r.get("CALIFICACION FINAL", 0)) for r in rows]
    total_aulas = len(finals)
    promedio = round(sum(finals) / total_aulas, 2) if total_aulas else 0.0

    cats = [final_qual(x)[1] for x in finals] if total_aulas else []
    exc = cats.count("EXCELENTE")
    bueno = cats.count("BUENO")
    acept = cats.count("ACEPTABLE")
    insat = cats.count("INSATISFACTORIO")

    def _pct(n):
        return round((n * 100.0) / total_aulas, 1) if total_aulas else 0.0

    pct_exc = _pct(exc)
    pct_bueno = _pct(bueno)
    pct_acept = _pct(acept)
    pct_insat = _pct(insat)

    kpi_cards = f"""
    <div style="display:flex;flex-wrap:wrap;gap:12px;margin:10px 0 8px 0;">
      <div style="flex:1 1 160px;background:#ffffff;border:1px solid {BRAND['table_border']};border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#667;letter-spacing:.4px;text-transform:uppercase;">Aulas revisadas</div>
        <div style="font-size:22px;font-weight:800;color:{BRAND['primary']};line-height:1.2;">{total_aulas}</div>
      </div>
      <div style="flex:1 1 160px;background:#ffffff;border:1px solid {BRAND['table_border']};border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#667;letter-spacing:.4px;text-transform:uppercase;">Promedio final</div>
        <div style="font-size:22px;font-weight:800;color:{BRAND['primary']};line-height:1.2;">{promedio}</div>
        <div style="font-size:11px;color:#667;">(0–100)</div>
      </div>
      <div style="flex:1 1 160px;background:#dcfce7;border:1px solid #cfead7;border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#14532d;letter-spacing:.4px;text-transform:uppercase;">Excelente</div>
        <div style="font-size:20px;font-weight:800;color:#14532d;line-height:1.2;">{exc}</div>
        <div style="font-size:11px;color:#14532d;">{pct_exc}%</div>
      </div>
      <div style="flex:1 1 160px;background:#dbeafe;border:1px solid #c3dafe;border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#1d4ed8;letter-spacing:.4px;text-transform:uppercase;">Bueno</div>
        <div style="font-size:20px;font-weight:800;color:#1d4ed8;line-height:1.2;">{bueno}</div>
        <div style="font-size:11px;color:#1d4ed8;">{pct_bueno}%</div>
      </div>
      <div style="flex:1 1 160px;background:#ffedd5;border:1px solid #f3c9a0;border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#92400e;letter-spacing:.4px;text-transform:uppercase;">Aceptable</div>
        <div style="font-size:20px;font-weight:800;color:#92400e;line-height:1.2;">{acept}</div>
        <div style="font-size:11px;color:#92400e;">{pct_acept}%</div>
      </div>
      <div style="flex:1 1 160px;background:#fee2e2;border:1px solid #f3c9cf;border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#7f1d1d;letter-spacing:.4px;text-transform:uppercase;">Insatisfactorio</div>
        <div style="font-size:20px;font-weight:800;color:#7f1d1d;line-height:1.2;">{insat}</div>
        <div style="font-size:11px;color:#7f1d1d;">{pct_insat}%</div>
      </div>
    </div>"""

    # Barra horizontal apilada
    if total_aulas > 0:
        w_exc = exc * 100.0 / total_aulas
        w_bueno = bueno * 100.0 / total_aulas
        w_acept = acept * 100.0 / total_aulas
        w_insat = max(0.0, 100.0 - (w_exc + w_bueno + w_acept))
    else:
        w_exc = w_bueno = w_acept = w_insat = 0.0

    bar_html = f"""
    <div style="width:100%;height:18px;border-radius:12px;overflow:hidden;border:1px solid {BRAND['table_border']};background:#f9fafb;margin-top:4px;white-space:nowrap;font-size:0;">
      <span style="display:inline-block;width:{w_exc:.4f}%;height:18px;background:#16a34a;"></span>
      <span style="display:inline-block;width:{w_bueno:.4f}%;height:18px;background:#2563eb;"></span>
      <span style="display:inline-block;width:{w_acept:.4f}%;height:18px;background:#ea580c;"></span>
      <span style="display:inline-block;width:{w_insat:.4f}%;height:18px;background:#b91c1c;"></span>
    </div>
    <div style="font-size:11px;color:#555;margin-top:3px;">
      Excelente: {exc} ({pct_exc}%) ·
      Bueno: {bueno} ({pct_bueno}%) ·
      Aceptable: {acept} ({pct_acept}%) ·
      Insatisf.: {insat} ({pct_insat}%)
    </div>"""

    legend = """
    <div style="font-size:11px;color:#444;margin:4px 0 2px 0;">
      <span style="display:inline-block;margin-right:10px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:#16a34a;margin-right:3px;"></span>
        Excelente
      </span>
      <span style="display:inline-block;margin-right:10px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:#2563eb;margin-right:3px;"></span>
        Bueno
      </span>
      <span style="display:inline-block;margin-right:10px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:#ea580c;margin-right:3px;"></span>
        Aceptable
      </span>
      <span style="display:inline-block;margin-right:10px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:#b91c1c;margin-right:3px;"></span>
        Insatisfactorio
      </span>
    </div>
    """

    resumen_block = f"""
    <div style="margin:8px 0 14px 0;padding:14px 16px;background:#f8fafc;border-radius:12px;border:1px solid {BRAND['table_border']};">
      <div style="font-size:13px;font-weight:600;color:{BRAND['primary_dark']};margin-bottom:4px;">
        Resumen de desempeño de sus aulas (Momento 2)
      </div>
      {kpi_cards}
      <div style="margin-top:6px;">
        {legend}
        {bar_html}
      </div>
    </div>
    """

    body = (
        saludo_docente(nombre, docente_id) +
        "<p>Desde el <strong>Campus Virtual</strong> realizamos el seguimiento de sus aulas "
        "en dos fases: <strong>Alistamiento</strong> y <strong>Ejecución</strong>. "
        "A continuación encontrará el resumen final de cada aula (nota de alistamiento, nota de ejecución y calificación final).</p>"
        + resumen_block +
        tabla_docente(rows) +
        "<div style='height:12px;'></div>" +
        bloque_mensaje_final_docente(rows) +
        "<p><strong>Contacto:</strong><br>"
        "Profesional de Campus Virtual: Jaime Duván Lozano Ardila<br>"
        "Correo: <a href='mailto:jaime.lozano.a@uniminuto.edu' style='color:#003366;text-decoration:underline;'>jaime.lozano.a@uniminuto.edu</a><br>"
        "Ubicación: Biblioteca – Sede Principal Chicalá</p>"
        + boton_agendar() +
        "<div style='height:18px;'></div>"
        "<div style='text-align:center;color:#333;font-size:14px;'>Campus Virtual – Rectoría Centro Sur</div>"
    )
    title = "Informe final de seguimiento – <span style='color:#f5b301;'>Campus Virtual RCS</span>"
    return email_shell(title, body)
def html_programa_resumen(programa, df_prog, col_docente_nm, col_docente_id):
    """
    Informe por programa (correo a coordinador).
    Incluye:
      - KPIs del programa
      - Una barra horizontal apilada (distribución excelente/bueno/aceptable/insatisfactorio)
      - Tabla de docentes con nº de aulas y promedio final
    Estilo unificado con el informe global y el informe de docentes.
    """
    dfp = df_prog.copy()
    total_aulas = int(len(dfp))
    finals = dfp["CALIFICACION FINAL"].astype(float)
    promedio = round(float(finals.mean()), 2) if total_aulas > 0 else 0.0

    # Distribución por desempeño
    def _cat(x):
        return final_qual(x)[1]

    cats = finals.apply(_cat) if total_aulas else []
    exc = int((cats == "EXCELENTE").sum()) if total_aulas else 0
    bueno = int((cats == "BUENO").sum()) if total_aulas else 0
    acept = int((cats == "ACEPTABLE").sum()) if total_aulas else 0
    insat = int((cats == "INSATISFACTORIO").sum()) if total_aulas else 0

    def _pct(n):
        return round((n * 100.0) / total_aulas, 1) if total_aulas else 0.0

    pct_exc = _pct(exc)
    pct_bueno = _pct(bueno)
    pct_acept = _pct(acept)
    pct_insat = _pct(insat)

    # ---- KPIs estilo informe global ----
    kpi_cards = f"""
    <div style="display:flex;flex-wrap:wrap;gap:12px;margin:8px 0 10px 0;">
      <div style="flex:1 1 180px;background:#ffffff;border:1px solid {BRAND['table_border']};border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#667;letter-spacing:.4px;text-transform:uppercase;">Aulas del programa</div>
        <div style="font-size:22px;font-weight:800;color:{BRAND['primary']};line-height:1.2;">{total_aulas}</div>
      </div>
      <div style="flex:1 1 180px;background:#ffffff;border:1px solid {BRAND['table_border']};border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#667;letter-spacing:.4px;text-transform:uppercase;">Promedio final</div>
        <div style="font-size:22px;font-weight:800;color:{BRAND['primary']};line-height:1.2;">{promedio}</div>
        <div style="font-size:11px;color:#667;">(0–100)</div>
      </div>
      <div style="flex:1 1 180px;background:#dcfce7;border:1px solid #cfead7;border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#14532d;letter-spacing:.4px;text-transform:uppercase;">Excelente</div>
        <div style="font-size:20px;font-weight:800;color:#14532d;line-height:1.2;">{exc}</div>
        <div style="font-size:11px;color:#14532d;">{pct_exc}%</div>
      </div>
      <div style="flex:1 1 180px;background:#dbeafe;border:1px solid #c3dafe;border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#1d4ed8;letter-spacing:.4px;text-transform:uppercase;">Bueno</div>
        <div style="font-size:20px;font-weight:800;color:#1d4ed8;line-height:1.2;">{bueno}</div>
        <div style="font-size:11px;color:#1d4ed8;">{pct_bueno}%</div>
      </div>
      <div style="flex:1 1 180px;background:#ffedd5;border:1px solid #f3c9a0;border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#92400e;letter-spacing:.4px;text-transform:uppercase;">Aceptable</div>
        <div style="font-size:20px;font-weight:800;color:#92400e;line-height:1.2;">{acept}</div>
        <div style="font-size:11px;color:#92400e;">{pct_acept}%</div>
      </div>
      <div style="flex:1 1 180px;background:#fee2e2;border:1px solid #f3c9cf;border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px rgba(0,0,0,.04);">
        <div style="font-size:11px;color:#7f1d1d;letter-spacing:.4px;text-transform:uppercase;">Insatisfactorio</div>
        <div style="font-size:20px;font-weight:800;color:#7f1d1d;line-height:1.2;">{insat}</div>
        <div style="font-size:11px;color:#7f1d1d;">{pct_insat}%</div>
      </div>
    </div>"""

    # ---- Barra horizontal apilada (solo este programa) ----
    if total_aulas > 0:
        w_exc = exc * 100.0 / total_aulas
        w_bueno = bueno * 100.0 / total_aulas
        w_acept = acept * 100.0 / total_aulas
        w_insat = max(0.0, 100.0 - (w_exc + w_bueno + w_acept))
    else:
        w_exc = w_bueno = w_acept = w_insat = 0.0

    bar_html = f"""
    <div style="width:100%;height:18px;border-radius:12px;overflow:hidden;border:1px solid {BRAND['table_border']};background:#f9fafb;margin-top:4px;white-space:nowrap;font-size:0;">
      <span style="display:inline-block;width:{w_exc:.4f}%;height:18px;background:#16a34a;"></span>
      <span style="display:inline-block;width:{w_bueno:.4f}%;height:18px;background:#2563eb;"></span>
      <span style="display:inline-block;width:{w_acept:.4f}%;height:18px;background:#ea580c;"></span>
      <span style="display:inline-block;width:{w_insat:.4f}%;height:18px;background:#b91c1c;"></span>
    </div>
    <div style="font-size:11px;color:#555;margin-top:3px;">
      Excelente: {exc} ({pct_exc}%) ·
      Bueno: {bueno} ({pct_bueno}%) ·
      Aceptable: {acept} ({pct_acept}%) ·
      Insatisf.: {insat} ({pct_insat}%)
    </div>"""

    legend = """
    <div style="font-size:11px;color:#444;margin:4px 0 2px 0;">
      <span style="display:inline-block;margin-right:10px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:#16a34a;margin-right:3px;"></span>
        Excelente
      </span>
      <span style="display:inline-block;margin-right:10px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:#2563eb;margin-right:3px;"></span>
        Bueno
      </span>
      <span style="display:inline-block;margin-right:10px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:#ea580c;margin-right:3px;"></span>
        Aceptable
      </span>
      <span style="display:inline-block;margin-right:10px;">
        <span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:#b91c1c;margin-right:3px;"></span>
        Insatisfactorio
      </span>
    </div>
    """

    # ---- Tabla de docentes (nº aulas y promedio) ----
    docentes = []
    for docente_id_val, gdoc in dfp.groupby(col_docente_id):
        nombre = next((str(x).strip() for x in gdoc[col_docente_nm].dropna().unique() if str(x).strip()), "")
        nombre = nombre or f"ID {to_int_or_str(docente_id_val)}"
        finals_doc = gdoc["CALIFICACION FINAL"].astype(float)
        prom_doc = round(float(finals_doc.mean()), 2) if len(gdoc) > 0 else 0.0
        docentes.append({
            "id": docente_id_val,
            "nombre": nombre,
            "num_aulas": len(gdoc),
            "promedio": prom_doc
        })
    docentes.sort(key=lambda d: str(d["nombre"]).upper())

    filas = []
    for i, d in enumerate(docentes):
        desc, short, fg, bg = final_qual(d["promedio"])
        badge = (
            f"<span class='badge-pill' "
            f"style='background:{fg};color:#fff;"
            f"border-radius:999px;padding:4px 10px;"
            f"font-weight:700;display:inline-block;"
            f"font-size:12px;line-height:1.3;min-width:72px;"
            f"text-align:center;'>"
            f"{d['promedio']} – {short.title()}</span>"
        )

        filas.append(f"""
<tr style="background:{BRAND['zebra'][i%2]};">
  <td style="padding:10px 12px;font-weight:600;color:{BRAND['primary_dark']};">{d['nombre']}</td>
  <td style="padding:10px 12px;text-align:center;">{to_int_or_str(d['id'])}</td>
  <td style="padding:10px 12px;text-align:center;">{d['num_aulas']}</td>
  <td style="padding:10px 12px;text-align:center;">{badge}</td>
</tr>""")
    html_tabla = "".join(filas)

    tabla_html = f"""
    <table width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;border-radius:8px;overflow:hidden;font-family:Segoe UI,Arial;table-layout:fixed;border:1px solid {BRAND['table_border']};margin-top:12px;">
      <thead class="thead-th">
        <tr style="background:{BRAND['primary_dark']};color:#fff;">
          <th style="padding:10px 12px;text-align:left;color:#fff!important;width:44%;">Docente</th>
          <th style="padding:10px 12px;text-align:center;color:#fff!important;width:18%;">ID</th>
          <th style="padding:10px 12px;text-align:center;color:#fff!important;width:18%;">Nº de aulas</th>
          <th style="padding:10px 12px;text-align:center;color:#fff!important;width:20%;">Promedio final</th>
        </tr>
      </thead>
      <tbody>{html_tabla}</tbody>
    </table>
    <div style="color:#667;margin-top:8px;font-size:12px;">
      {leyenda_html_final()}.<br>
      <em>El PDF adjunto contiene el detalle por NRC (Alistamiento y Ejecución) de cada aula.</em>
    </div>
    """

    shell = f"""
<p style="margin:0 0 6px 0;"><strong>Programa:</strong> {programa}</p>
<p style="margin:0 0 8px 0;">
  Este informe presenta el <strong>resultado final del Momento 2</strong> (Alistamiento + Ejecución)
  para las aulas del programa. A continuación se resumen los indicadores generales
  y el desempeño promedio por docente.
</p>
<div style="margin:8px 0 12px 0;padding:14px 16px;background:#f8fafc;border-radius:12px;border:1px solid {BRAND['table_border']};">
  {kpi_cards}
  <div style="margin-top:6px;">
    <div style="font-size:13px;font-weight:600;color:{BRAND['primary_dark']};margin-bottom:2px;">
      Distribución del desempeño final de las aulas del programa
    </div>
    {legend}
    {bar_html}
  </div>
</div>
{tabla_html}
"""
    title = f"Informe final – Programa <span style='color:#FFD000;'>{programa}</span>"
    return email_shell(title, shell)
def html_programa_detalle_global(programa, df_prog, col_docente_nm, col_docente_id):
    bloques = []
    for docente_id_val, gdoc in df_prog.groupby(col_docente_id):
        nombre = next((str(x).strip() for x in gdoc[col_docente_nm].dropna().unique() if str(x).strip()), "")
        nombre = nombre or f"ID {to_int_or_str(docente_id_val)}"

        filas = []
        for _, r in gdoc.sort_values(by=["NRC"]).iterrows():
            fase1 = r.get("CALIFICACION", 0)
            fase2 = r.get("CALIFICACION 2", 0)
            final = r.get("CALIFICACION FINAL", 0)
            desc, short, fg, bg = final_qual(final)
            puntajes_html = (
                f"Alistamiento: {to_int_or_str(fase1)}<br>"
                f"Ejecución: {to_int_or_str(fase2)}<br>"
                f"<strong>Final: {to_int_or_str(final)}</strong>"
            )
            rev = observacion_badge(r.get("OBSERVACION", ""))

            filas.append(f"""
            <tr style="background:{BRAND['zebra'][0]};font-size:13px;">
              <td style="padding:10px;width:14%;text-align:center;">{r.get('NRC','')}</td>
              <td style="padding:10px;width:50%;text-align:left;word-break:break-word;white-space:normal;overflow-wrap:anywhere;">{r.get('ASIGNATURA','')}</td>
              <td style="padding:10px;width:18%;text-align:left;font-size:12px;line-height:1.4;">{puntajes_html}</td>
              <td style="padding:10px;width:10%;text-align:center;">
                <span style="display:inline-block;padding:4px 10px;border-radius:999px;background:{fg};color:#fff;font-size:12px;font-weight:600;" class="badge-pill">
                  {short.title()}
                </span>
              </td>
              <td style="padding:10px;width:14%;text-align:left;white-space:normal;word-break:break-word;overflow-wrap:anywhere;line-height:1.35;">{rev}</td>
            </tr>""")

        tabla = f"""
        <table width="100%" style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px;border:1px solid #d9e0ef;table-layout:fixed;">
          <thead class="thead-th">
            <tr style="background:{BRAND['primary_dark']};color:#fff;">
              <th style="padding:10px;text-align:center;width:14%;color:#fff!important;">NRC</th>
              <th style="padding:10px;text-align:left;width:50%;color:#fff!important;">Asignatura</th>
              <th style="padding:10px;text-align:left;width:18%;color:#fff!important;">Puntajes (Fase 1 y 2)</th>
              <th style="padding:10px;text-align:center;width:10%;color:#fff!important;">Desempeño</th>
              <th style="padding:10px;text-align:center;width:14%;color:#fff!important;">Revisión</th>
            </tr>
          </thead>
          <tbody>{''.join(filas)}</tbody>
        </table>"""

        bloque = f"""
        <div style="margin:18px 0;padding:10px 14px;background:#fefefe;border:1px solid {BRAND['table_border']};border-left:6px solid {BRAND['accent']};border-radius:10px;">
          <div style="font-size:15px;color:{BRAND['primary_dark']};font-weight:700;margin-bottom:6px;">
            {nombre} <span style="font-weight:400;color:#667;">(ID: {to_int_or_str(docente_id_val)})</span>
          </div>
          {tabla}
        </div>"""
        bloques.append(bloque)

    wrapper = f"""
<div style="font-family:Segoe UI, Arial, sans-serif;max-width:860px;margin:0 auto;">
  <div style="background:{BRAND['primary']};color:#fff;padding:16px 20px;border-radius:10px 10px 0 0;border:1px solid #002b55;">
    <div style="font-size:20px;font-weight:700;">Detalle final por NRC – Programa <span style="color:#FFD000;">{programa}</span></div>
    <div style="font-size:12px;font-weight:400;margin-top:6px;color:#e6eaf2;">Momento 2 – Informe final (Alistamiento + Ejecución).</div>
  </div>
  <div style="border:1px solid {BRAND['table_border']};border-top:none;border-radius:0 0 10px 10px;padding:20px;background:{BRAND['panel_bg']};">
    {''.join(bloques)}
    <div style="margin-top:10px;color:#555;font-size:12px;text-align:right;">{leyenda_html_final()}</div>
  </div>
</div>"""
    return wrapper


def html_programa_detalle_mail(programa, df_prog, col_docente_nm, col_docente_id):
    cuerpo = html_programa_detalle_global(programa, df_prog, col_docente_nm, col_docente_id)
    title = f"Informe final – Programa <span style='color:#FFD000;'>{programa}</span>"
    mensaje = ("<p style='margin:0 0 12px 0;'>A continuación se presenta el "
               "<strong>detalle final por NRC</strong> del programa, con las notas de "
               "alistamiento, ejecución y su calificación final.</p>")
    return email_shell(title, mensaje + f"<div>{cuerpo}</div>")


# ---------- GLOBAL ----------

def build_program_stats(df, col_prog, col_puntaje_final):
    stats = []
    finals_all = df[col_puntaje_final].astype(float)

    for programa, g in df.groupby(col_prog):
        finals = g[col_puntaje_final].astype(float)
        aulas_total = len(g)
        promedio_programa = round(finals.mean(), 2) if aulas_total > 0 else 0.0

        def _cat(x):
            return final_qual(x)[1]

        cats = finals.apply(_cat)
        exc = int((cats == "EXCELENTE").sum())
        bueno = int((cats == "BUENO").sum())
        acept = int((cats == "ACEPTABLE").sum())
        insat = int((cats == "INSATISFACTORIO").sum())

        stats.append({
            "programa": programa,
            "aulas_total": aulas_total,
            "promedio": promedio_programa,
            "exc": exc,
            "bueno": bueno,
            "acept": acept,
            "insat": insat,
        })
    stats.sort(key=lambda x: str(x["programa"]).upper())
    return stats


def build_overall_totals(df, col_puntaje_final):
    finals = df[col_puntaje_final].astype(float)
    aulas_total = len(df)
    promedio_global = round(finals.mean(), 2) if aulas_total > 0 else 0.0

    def _cat(x):
        return final_qual(x)[1]

    cats = finals.apply(_cat)
    exc = int((cats == "EXCELENTE").sum())
    bueno = int((cats == "BUENO").sum())
    acept = int((cats == "ACEPTABLE").sum())
    insat = int((cats == "INSATISFACTORIO").sum())

    def pct(x):
        return round((x / aulas_total) * 100, 1) if aulas_total else 0.0

    return {
        "aulas_total": aulas_total,
        "promedio": promedio_global,
        "exc": exc,
        "pct_exc": pct(exc),
        "bueno": bueno,
        "pct_bueno": pct(bueno),
        "acept": acept,
        "pct_acept": pct(acept),
        "insat": insat,
        "pct_insat": pct(insat),
    }

def html_global_program_bars(df, col_prog, col_puntaje_final):
    """
    Bloque de barras horizontales apiladas (100%) por programa académico.
    Cada barra muestra la distribución de aulas en:
    Excelente, Bueno, Aceptable e Insatisfactorio.
    Se usa el mismo criterio de desempeño que en el resto del informe.
    """
    stats = build_program_stats(df, col_prog, col_puntaje_final)

    # Ordenamos de mayor a menor número de aulas para que la gráfica sea más clara
    stats = sorted(stats, key=lambda x: x["aulas_total"], reverse=True)

    filas = []
    for st in stats:
        total = st["aulas_total"] or 1  # evitar división por cero
        exc = st["exc"]
        bueno = st["bueno"]
        acept = st["acept"]
        insat = st["insat"]

        # porcentajes (por si quieres mostrarlos luego)
        pct_exc = round(exc * 100 / total, 1)
        pct_bueno = round(bueno * 100 / total, 1)
        pct_acept = round(acept * 100 / total, 1)
        pct_insat = round(insat * 100 / total, 1)

        filas.append(f"""
        <div style="display:flex;align-items:center;margin:6px 0;">
          <div style="width:30%;min-width:170px;padding-right:10px;font-size:13px;color:{BRAND['primary_dark']};font-weight:600;word-break:break-word;">
            {st['programa']}<br>
            <span style="font-size:11px;color:#667;font-weight:400;">
              Aulas: {total} · Prom: {st['promedio']}
            </span>
          </div>
          <div style="flex:1;display:flex;flex-direction:column;gap:4px;">
            <div style="display:flex;display:-webkit-box;display:-ms-flexbox;height:18px;border-radius:999px;overflow:hidden;border:1px solid {BRAND['table_border']};background:#f9fafb;">
              <div style="flex:{exc} 1 0%;-webkit-box-flex:{exc};-ms-flex:{exc} 1 0%;width:{pct_exc}%;background:#16a34a;font-size:0;"></div>
              <div style="flex:{bueno} 1 0%;-webkit-box-flex:{bueno};-ms-flex:{bueno} 1 0%;width:{pct_bueno}%;background:#2563eb;font-size:0;"></div>
              <div style="flex:{acept} 1 0%;-webkit-box-flex:{acept};-ms-flex:{acept} 1 0%;width:{pct_acept}%;background:#ea580c;font-size:0;"></div>
              <div style="flex:{insat} 1 0%;-webkit-box-flex:{insat};-ms-flex:{insat} 1 0%;width:{pct_insat}%;background:#b91c1c;font-size:0;"></div>
            </div>
            <div style="font-size:11px;color:#555;">
              Excelente: {exc} ({pct_exc}%) ·
              Bueno: {bueno} ({pct_bueno}%) ·
              Aceptable: {acept} ({pct_acept}%) ·
              Insatisf.: {insat} ({pct_insat}%)
            </div>
          </div>
        </div>""")

    if not filas:
        return ""

    return f"""
    <div style="margin:8px 0 16px 0;">
      <div style="font-size:15px;font-weight:600;color:{BRAND['primary_dark']};margin-bottom:8px;">
        Desempeño por programa académico
      </div>
      <div>
        {''.join(filas)}
      </div>
      <div style="font-size:11px;color:#666;margin-top:6px;">
        Cada barra representa el 100% de las aulas del programa, segmentadas por nivel de desempeño final
        (excelente, bueno, aceptable e insatisfactorio) según la calificación final.
      </div>
    </div>"""

def html_global_summary_table(df, col_prog, col_puntaje_final):
    stats = build_program_stats(df, col_prog, col_puntaje_final)
    tot = build_overall_totals(df, col_puntaje_final)

    # Tarjetas KPI superiores (números globales)
    kpi_cards = f"""
    <div style="display:flex;flex-wrap:wrap;gap:12px;margin:12px 0 16px 0;">
      <div style="flex:1 1 180px;background:#ffffff;border:1px solid {BRAND['table_border']};border-radius:12px;padding:14px 16px;box-shadow:0 2px 6px rgba(0,0,0,.05);">
        <div style="font-size:12px;color:#667;letter-spacing:.4px;text-transform:uppercase;">Aulas total</div>
        <div style="font-size:28px;font-weight:800;color:{BRAND['primary']};line-height:1.2;">{tot['aulas_total']}</div>
      </div>
      <div style="flex:1 1 180px;background:#ffffff;border:1px solid {BRAND['table_border']};border-radius:12px;padding:14px 16px;box-shadow:0 2px 6px rgba(0,0,0,.05);">
        <div style="font-size:12px;color:#667;letter-spacing:.4px;text-transform:uppercase;">Promedio final</div>
        <div style="font-size:28px;font-weight:800;color:{BRAND['primary']};line-height:1.2;">{tot['promedio']}</div>
        <div style="font-size:12px;color:#667;">(0–100)</div>
      </div>
      <div style="flex:1 1 180px;background:#dcfce7;border:1px solid #cfead7;border-radius:12px;padding:14px 16px;box-shadow:0 2px 6px rgba(0,0,0,.05);">
        <div style="font-size:12px;color:#14532d;letter-spacing:.4px;text-transform:uppercase;">Excelente</div>
        <div style="font-size:28px;font-weight:800;color:#14532d;line-height:1.2;">{tot['exc']}</div>
        <div style="font-size:12px;color:#14532d;">{tot['pct_exc']}%</div>
      </div>
      <div style="flex:1 1 180px;background:#dbeafe;border:1px solid #c3dafe;border-radius:12px;padding:14px 16px;box-shadow:0 2px 6px rgba(0,0,0,.05);">
        <div style="font-size:12px;color:#1d4ed8;letter-spacing:.4px;text-transform:uppercase;">Bueno</div>
        <div style="font-size:28px;font-weight:800;color:#1d4ed8;line-height:1.2;">{tot['bueno']}</div>
        <div style="font-size:12px;color:#1d4ed8;">{tot['pct_bueno']}%</div>
      </div>
      <div style="flex:1 1 180px;background:#ffedd5;border:1px solid #f3c9a0;border-radius:12px;padding:14px 16px;box-shadow:0 2px 6px rgba(0,0,0,.05);">
        <div style="font-size:12px;color:#92400e;letter-spacing:.4px;text-transform:uppercase;">Aceptable</div>
        <div style="font-size:28px;font-weight:800;color:#92400e;line-height:1.2;">{tot['acept']}</div>
        <div style="font-size:12px;color:#92400e;">{tot['pct_acept']}%</div>
      </div>
      <div style="flex:1 1 180px;background:#fee2e2;border:1px solid #f3c9cf;border-radius:12px;padding:14px 16px;box-shadow:0 2px 6px rgba(0,0,0,.05);">
        <div style="font-size:12px;color:#7f1d1d;letter-spacing:.4px;text-transform:uppercase;">Insatisfactorio</div>
        <div style="font-size:28px;font-weight:800;color:#7f1d1d;line-height:1.2;">{tot['insat']}</div>
        <div style="font-size:12px;color:#7f1d1d;">{tot['pct_insat']}%</div>
      </div>
    </div>"""

    # 🔹 NUEVA: gráfica horizontal por programas, debajo de los KPI
    bars_block = html_global_program_bars(df, col_prog, col_puntaje_final)

    filas = []
    for i, row in enumerate(stats):
        filas.append(f"""
        <tr style="background:{BRAND['zebra'][i%2]};font-size:14px;">
          <td style="padding:8px 12px;text-align:left;font-weight:600;color:{BRAND['primary_dark']};">{row['programa']}</td>
          <td style="padding:8px 12px;text-align:center;">{row['aulas_total']}</td>
          <td style="padding:8px 12px;text-align:center;">{row['promedio']}</td>
          <td style="padding:8px 12px;text-align:center;background:#dcfce7;">{row['exc']}</td>
          <td style="padding:8px 12px;text-align:center;background:#dbeafe;">{row['bueno']}</td>
          <td style="padding:8px 12px;text-align:center;background:#ffedd5;">{row['acept']}</td>
          <td style="padding:8px 12px;text-align:center;background:#fee2e2;">{row['insat']}</td>
        </tr>""")

    filas.append(f"""
    <tr style="background:#FFF7D6;font-size:14px;border-top:2px solid #e3e8f1;">
      <td style="padding:10px 12px;text-align:left;font-weight:800;color:#6b4d00;">TOTAL RECTORÍA</td>
      <td style="padding:10px 12px;text-align:center;font-weight:700;color:#6b4d00;">{tot['aulas_total']}</td>
      <td style="padding:10px 12px;text-align:center;font-weight:700;color:#6b4d00;">{tot['promedio']}</td>
      <td style="padding:10px 12px;text-align:center;background:#dcfce7;font-weight:700;color:#14532d;">{tot['exc']}</td>
      <td style="padding:10px 12px;text-align:center;background:#dbeafe;font-weight:700;color:#1d4ed8;">{tot['bueno']}</td>
      <td style="padding:10px 12px;text-align:center;background:#ffedd5;font-weight:700;color:#92400e;">{tot['acept']}</td>
      <td style="padding:10px 12px;text-align:center;background:#fee2e2;font-weight:700;color:#7f1d1d;">{tot['insat']}</td>
    </tr>""")

    tabla = f"""
    <table width="100%" cellspacing="0" cellpadding="0" border="0"
           style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:14px;border-radius:8px;overflow:hidden;table-layout:fixed;border:1px solid {BRAND['table_border']}">
      <thead class="thead-th">
        <tr style="background:{BRAND['primary']};color:#fff;">
          <th style="padding:10px 12px;text-align:left;color:#fff!important;width:32%;">Programa</th>
          <th style="padding:10px 12px;text-align:center;color:#fff!important;width:10%;">Aulas</th>
          <th style="padding:10px 12px;text-align:center;color:#fff!important;width:10%;">Promedio</th>
          <th style="padding:10px 12px;text-align:center;color:#fff!important;width:12%;">Excelente</th>
          <th style="padding:10px 12px;text-align:center;color:#fff!important;width:12%;">Bueno</th>
          <th style="padding:10px 12px;text-align:center;color:#fff!important;width:12%;">Aceptable</th>
          <th style="padding:10px 12px;text-align:center;color:#fff!important;width:12%;">Insatisf.</th>
        </tr>
      </thead>
      <tbody>{''.join(filas)}</tbody>
    </table>
    <div style="font-size:12px;color:#666;margin-top:8px;">
      La clasificación se basa en la calificación final (0–100): excelente (91–100), bueno (80–90),
      aceptable (70–79) e insatisfactorio (0–69).
    </div>
    """

    header_card = f"""
    <div style="background:{BRAND['primary']};color:#fff;padding:16px 20px;border-radius:10px 10px 0 0;border:1px solid #002b55;">
      <div style="font-size:20px;font-weight:700;">Informe global – Programas académicos (Rectoría Centro Sur)</div>
      <div style="font-size:12px;font-weight:400;margin-top:6px;color:#e6eaf2;">
        Momento 2 – Informe final (Fase de Alistamiento 50% + Fase de Ejecución 50%).
      </div>
    </div>"""

    cuerpo_card = f"""
    <div style="border:1px solid {BRAND['table_border']};border-top:none;border-radius:0 0 10px 10px;padding:20px;background:{BRAND['panel_bg']};">
      {kpi_cards}
      {bars_block}
      <div style="font-size:15px;color:{BRAND['primary_dark']};font-weight:600;margin:8px 0 12px 0;">
        Resumen consolidado por programa (desempeño final)
      </div>
      {tabla}
    </div>"""

    return f"<div style='max-width:980px;margin:0 auto 24px auto;font-family:Segoe UI,Arial,sans-serif;'>{header_card}{cuerpo_card}</div>"


def html_global_programas_resumen(df, col_prog, col_docente_nm, col_docente_id, col_puntaje_final):
    bloque_top = html_global_summary_table(df, col_prog, col_puntaje_final)
    bloques_programas = []
    for programa, gprog in df.groupby(col_prog):
        bloque = html_programa_resumen(programa, gprog, col_docente_nm, col_docente_id)
        bloques_programas.append(f"<div style='margin:18px auto;max-width:900px;'>{bloque}</div>")
    pagina = f"""
<div style="font-family:Segoe UI, Arial, sans-serif;">
  {bloque_top}
  <div style="max-width:980px;margin:0 auto;border:1px solid {BRAND['table_border']};border-radius:10px;background:#fff;box-shadow:0 4px 12px rgba(0,0,0,.05);padding:20px;">
    <div style="font-size:16px;font-weight:700;color:{BRAND['primary_dark']};margin-bottom:12px;">
      Detalle por programa (docentes, nº de aulas y promedio final)
    </div>
    {''.join(bloques_programas)}
    {footer_block()}
  </div>
</div>"""
    return pagina


# ---------- OUTLOOK ----------

def outlook_send(to_email, subject, html_body, attachments=None, cc=None, bcc=None, reply_to=None, dry_run=False):
    if dry_run:
        print(f"[DRY-RUN] To: {to_email} | Subject: {subject} | Adjuntos: {len(attachments or [])}")
        return
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_email
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    if reply_to:
        try:
            PR_REPLY_RECIPIENT_NAMES = "http://schemas.microsoft.com/mapi/proptag/0x0E04001E"
            mail.PropertyAccessor.SetProperty(PR_REPLY_RECIPIENT_NAMES, reply_to)
        except Exception:
            pass
    mail.Subject = subject
    mail.HTMLBody = html_body
    for att in resolve_existing_paths(attachments):
        try:
            mail.Attachments.Add(att)
        except Exception as e:
            print(f"⚠️ No se pudo adjuntar {att}: {e}")
    mail.Send()


# ---------- MAIN ----------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", required=True)
    parser.add_argument("--mode", choices=["preview", "outlook"], default="preview")
    parser.add_argument("--out", default="./salida")
    parser.add_argument("--send", default="docentes,programas")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--only-docentes")
    parser.add_argument("--only-correos")
    parser.add_argument("--only-programas")
    parser.add_argument("--only")
    parser.add_argument("--coords", help="coordinadores.csv (PROGRAMA,PROGRAMA_CORTO,COORDINADOR,EMAIL)")
    parser.add_argument("--attach-programa")
    parser.add_argument("--attach-docente")
    parser.add_argument("--limit-docentes", type=int)
    parser.add_argument("--limit-programas", type=int)
    parser.add_argument("--make-global", action="store_true")
    parser.add_argument("--send-global", action="store_true")
    parser.add_argument("--global-to")
    parser.add_argument("--force-to")
    parser.add_argument("--cc")
    parser.add_argument("--bcc")
    parser.add_argument("--reply-to")
    args = parser.parse_args()

    outdir = Path(args.out)
    (outdir / "docentes").mkdir(parents=True, exist_ok=True)
    (outdir / "programas").mkdir(parents=True, exist_ok=True)
    (outdir / "global").mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(args.excel)
    required_cols = [
        "PROGRAMA",
        "ID DOCENTE",
        "CORREO",
        "CALIFICACION",
        "CALIFICACION 2",
        "CALIFICACION FINAL",
        "NRC",
        "OBSERVACION",
    ]
    for col in required_cols:
        if col not in df.columns:
            raise SystemExit(f"Falta la columna requerida en el Excel: {col}")
    if len(df.columns) < 5:
        raise SystemExit("El Excel no tiene al menos 5 columnas para tomar el nombre del docente (columna E).")

    df = normalize_dataframe(df)
    col_docente_nm = df.columns[4]
    col_docente_id = "ID DOCENTE"
    col_correo = "CORREO"
    col_prog = "PROGRAMA"
    col_puntaje_final = "CALIFICACION FINAL"
    col_nrc = "NRC"
    col_observ = "OBSERVACION"
    col_asig = "ASIGNATURA" if "ASIGNATURA" in df.columns else None

    send_modes = [s.strip().lower() for s in args.send.split(",") if s.strip()]

    only_ids = set([s.strip() for s in str(args.only or "").split(",") if s.strip()])
    if args.only_docentes:
        only_ids |= set([s.strip() for s in str(args.only_docentes).split(",") if s.strip()])

    only_emails = set()
    if args.only_correos:
        sep = ';' if ';' in args.only_correos else ','
        only_emails = set([s.strip() for s in args.only_correos.split(sep) if s.strip()])

    only_programs = set([s.strip() for s in str(args.only_programas or "").split(",") if s.strip()])

    # --- coordinadores: lectura robusta ---
    coords_map = {}
    if args.coords and Path(args.coords).exists():
        cdf = pd.read_csv(
            args.coords,
            sep=None,
            engine="python",
            encoding="utf-8-sig"
        )
        import re as _re
        cdf.columns = [
            _re.sub(r"[\uFEFF\xa0]", "", str(c)).strip().upper()
            for c in cdf.columns
        ]
        cdf = cdf.rename(columns={
            "PROGRAMA ": "PROGRAMA",
            "PROGRAMA_CORTO ": "PROGRAMA_CORTO",
            "COORDINADOR ": "COORDINADOR",
            "EMAIL ": "EMAIL",
        })
        required_cols_coords = {"PROGRAMA", "PROGRAMA_CORTO", "COORDINADOR", "EMAIL"}
        missing = required_cols_coords - set(cdf.columns)
        if missing:
            raise SystemExit(
                f"coordinadores.csv no tiene columnas requeridas: {missing}. "
                f"Columnas leídas: {list(cdf.columns)}"
            )
        for _, row in cdf.iterrows():
            prog = str(row["PROGRAMA"]).strip()
            coords_map[prog] = {
                "corto": str(row.get("PROGRAMA_CORTO", "")).strip(),
                "coord": str(row.get("COORDINADOR", "")).strip(),
                "email": str(row.get("EMAIL", "")).strip(),
            }

    docente_extra_attachments = []
    if args.attach_docente:
        raw = args.attach_docente
        sep = ';' if ';' in raw else ','
        docente_extra_attachments = resolve_existing_paths([s.strip() for s in raw.split(sep) if s.strip()])
    else:
        default_path = Path.cwd() / DEFAULT_DOCENTE_ATTACH
        if default_path.exists():
            docente_extra_attachments = [str(default_path.resolve())]
        else:
            excel_dir = Path(args.excel).resolve().parent
            candidate = excel_dir / DEFAULT_DOCENTE_ATTACH
            if candidate.exists():
                docente_extra_attachments = [str(candidate.resolve())]

    # ----- DOCENTES -----
    if "docentes" in send_modes:
        count_doc = 0
        for docente_id_val, g in df.groupby(col_docente_id):
            if args.limit_docentes is not None and count_doc >= args.limit_docentes:
                break
            if only_programs:
                progs_doc = set(str(x).strip() for x in g[col_prog].unique())
                if progs_doc.isdisjoint(only_programs):
                    continue
            if only_ids:
                try:
                    this_id_str = str(int(float(docente_id_val)))
                except Exception:
                    this_id_str = str(docente_id_val)
                if this_id_str not in only_ids:
                    continue
            correo = next((str(x).strip() for x in g[col_correo].dropna().unique()
                           if str(x).strip() and is_email(str(x).strip())), None)
            if only_emails and (correo not in only_emails):
                continue
            nombre = next((str(x).strip() for x in g[col_docente_nm].dropna().unique()
                           if str(x).strip()), None)

            rows = []
            for _, r in g.iterrows():
                rows.append({
                    "NRC": r.get(col_nrc, ""),
                    "ASIGNATURA": r.get(col_asig, "") if col_asig else "",
                    "PROGRAMA": r.get(col_prog, ""),
                    "CALIFICACION": r.get("CALIFICACION", 0),
                    "CALIFICACION 2": r.get("CALIFICACION 2", 0),
                    "CALIFICACION FINAL": r.get("CALIFICACION FINAL", 0),
                    "OBSERVACION": r.get(col_observ, ""),
                })

            html = html_docente(nombre, docente_id_val, rows)
            fname = (nombre or str(docente_id_val) or "docente").replace(" ", "_").replace("/", "_")
            (outdir / "docentes" / f"{FECHA_ETQ}_docente_{fname}.html").write_text(html, encoding="utf-8")

            if args.mode == "outlook":
                to_email = args.force_to if args.force_to else correo
                if to_email and is_email(to_email):
                    subject = SUBJECT_DOCENTE.format(DOCENTE_LBL=(nombre or f"ID {to_int_or_str(docente_id_val)}"))
                    if args.force_to:
                        subject = f"[PRUEBA] {subject}"
                    attachments = docente_extra_attachments.copy()
                    try:
                        outlook_send(
                            to_email, subject, html,
                            attachments=attachments,
                            cc=None if not args.cc else "; ".join(parse_emails(args.cc)),
                            bcc=None if not args.bcc else "; ".join(parse_emails(args.bcc)),
                            reply_to=args.reply_to, dry_run=args.dry_run
                        )
                        log_envio(outdir / "envios.csv", "docente", to_email, subject, attachments)
                        print(f"✅ Docente enviado: {nombre} -> {to_email} (adjuntos: {len(attachments)})")
                    except Exception as e:
                        print(f"⚠️ Error enviando a {nombre}: {e}")
                else:
                    print(f"❌ {nombre} sin correo válido")
            count_doc += 1

    # ----- PROGRAMAS -----
    count_prog = 0
    if "programas" in send_modes:
        for programa, gprog in df.groupby(col_prog):
            if args.limit_programas is not None and count_prog >= args.limit_programas:
                break
            if only_programs and (str(programa).strip() not in only_programs):
                continue

            resumen_html = html_programa_resumen(programa, gprog, col_docente_nm, col_docente_id)
            fname_prog = str(programa).replace(" ", "_").replace("/", "_")
            (outdir / "programas" / f"{FECHA_ETQ}_{fname_prog}__resumen.html").write_text(resumen_html, encoding="utf-8")

            mail_html = html_programa_detalle_mail(programa, gprog, col_docente_nm, col_docente_id)
            detalle_html_puro = html_programa_detalle_global(programa, gprog, col_docente_nm, col_docente_id)
            detalle_html_path = (outdir / "programas" / f"{FECHA_ETQ}_{fname_prog}__detalle.html").resolve()
            detalle_html_path.write_text(detalle_html_puro, encoding="utf-8")

            attachments = []
            if PDFKIT_AVAILABLE and PDFKIT_CONFIG:
                try:
                    detalle_pdf_path = (outdir / "programas" /
                                        f"RCS_{FECHA_ETQ}_{fname_prog}__detalle.pdf").resolve()
                    pdfkit.from_string(
                        wrap_for_pdf(detalle_html_puro),
                        str(detalle_pdf_path),
                        configuration=PDFKIT_CONFIG,
                        options=PDF_OPTIONS
                    )
                    attachments.append(str(detalle_pdf_path))
                except Exception as e:
                    print(f"⚠️ No se pudo generar PDF para {programa}. Se adjunta HTML. {e}")
                    attachments.append(str(detalle_html_path))
            else:
                attachments.append(str(detalle_html_path))

            if args.attach_programa:
                raw = args.attach_programa
                sep = ';' if ';' in raw else ','
                attachments += [s.strip() for s in raw.split(sep) if s.strip()]

            if args.mode == "outlook":
                # Si hay --force-to SIEMPRE se usa (modo prueba)
                if args.force_to:
                    to_email = args.force_to
                elif args.coords and (programa in coords_map) and is_email(coords_map[programa]["email"]):
                    to_email = coords_map[programa]["email"]
                else:
                    to_email = None

                if to_email:
                    subject = SUBJECT_PROGRAMA.format(PROGRAMA=programa)
                    if args.force_to:
                        subject = f"[PRUEBA] {subject}"
                    try:
                        outlook_send(
                            to_email, subject, mail_html,
                            attachments=attachments,
                            cc=None if not args.cc else "; ".join(parse_emails(args.cc)),
                            bcc=None if not args.bcc else "; ".join(parse_emails(args.bcc)),
                            reply_to=args.reply_to, dry_run=args.dry_run
                        )
                        log_envio(outdir / "envios.csv", "programa", to_email, subject, attachments)
                        print(f"📨 Programa '{programa}' enviado a {to_email} (adjuntos: {len(attachments)})")
                    except Exception as e:
                        print(f"⚠️ Error enviando programa '{programa}' a {to_email}: {e}")
                else:
                    print(f"❌ Sin correo de coordinador para '{programa}' y sin --force-to. Solo generado HTML/PDF.")
            count_prog += 1

    # ----- GLOBAL -----
    if args.make_global:
        global_html = html_global_programas_resumen(df, col_prog, col_docente_nm, col_docente_id, col_puntaje_final)
        global_html_path = (outdir / "global" / "global_programas__resumen.html")
        global_html_path.write_text(global_html, encoding="utf-8")

        global_pdf_path = None
        if PDFKIT_AVAILABLE and PDFKIT_CONFIG:
            try:
                global_pdf_path = (outdir / "global" /
                                   f"RCS_{FECHA_ETQ}_global_programas__resumen.pdf").resolve()
                pdfkit.from_string(
                    wrap_for_pdf(global_html),
                    str(global_pdf_path),
                    configuration=PDFKIT_CONFIG,
                    options=PDF_OPTIONS
                )
                print(f"📄 Global PDF: {global_pdf_path}")
            except Exception as e:
                print(f"⚠️ No se pudo generar PDF global: {e}")

        if args.send_global and args.mode == "outlook":
            inner = global_html_path.read_text(encoding="utf-8")
            mail_body = email_shell("Informe global final – <span style='color:#FFD000;'>Campus Virtual RCS</span>", inner)
            recipients = [args.force_to] if args.force_to else parse_emails(args.global_to)
            if not recipients:
                print("⚠️ No hay destinatarios para el global. Usa --global-to o --force-to.")
            else:
                attachments = []
                if global_pdf_path and Path(global_pdf_path).exists():
                    attachments.append(str(global_pdf_path))
                else:
                    attachments.append(str(global_html_path.resolve()))
                try:
                    to_field = "; ".join(recipients)
                    outlook_send(
                        to_field, SUBJECT_GLOBAL, mail_body, attachments=attachments,
                        cc=None if not args.cc else "; ".join(parse_emails(args.cc)),
                        bcc=None if not args.bcc else "; ".join(parse_emails(args.bcc)),
                        reply_to=args.reply_to, dry_run=args.dry_run
                    )
                    log_envio(outdir / "envios.csv", "global", to_field, SUBJECT_GLOBAL, attachments)
                    print(f"📨 Global enviado a: {to_field} (adjuntos: {len(attachments)})")
                except Exception as e:
                    print(f"⚠️ Error enviando Global: {e}")

    print("Proceso finalizado ✅")
    print(f"HTML por docente:  {outdir / 'docentes'}")
    print(f"Programas (resumen/detalle): {outdir / 'programas'}")
    print(f"Global: {outdir / 'global'}")


if __name__ == "__main__":
    main()
