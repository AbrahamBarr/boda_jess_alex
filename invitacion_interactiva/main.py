from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
from datetime import datetime
import os
import json
import io
import pytz
from google.oauth2 import service_account
from googleapiclient.discovery import build


os.makedirs("data", exist_ok=True)  

TIMEZONE = pytz.timezone("America/Mexico_City")
SHEET_ID = os.getenv("SHEET_ID")                      # requerido
SHEET_RANGE = os.getenv("SHEET_RANGE", "Respuestas!A:C")  # A:Nombre, B:Asistentes, C:Fecha Confirmación

def _sheets_service():
    """Crea el cliente de Google Sheets desde la variable de entorno GCP_SA_KEY."""
    gcp_sa_key = os.getenv("GCP_SA_KEY")
    if not gcp_sa_key:
        raise RuntimeError("Falta la variable de entorno GCP_SA_KEY")
    try:
        key_dict = json.loads(gcp_sa_key)
    except json.JSONDecodeError as e:
        raise RuntimeError("GCP_SA_KEY no es JSON válido. Verifica que pegaste el JSON completo.") from e

    creds = service_account.Credentials.from_service_account_info(
        key_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds)

def append_to_sheet_row(nombre: str, asistentes: int, fecha_str: str):
    """Agrega una fila al Sheet en el orden: [Nombre, Asistentes, Fecha Confirmación]."""
    if not SHEET_ID:
        raise RuntimeError("Falta SHEET_ID")
    service = _sheets_service()
    body = {"values": [[nombre, asistentes, fecha_str]]}
    return service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body
    ).execute()

def read_sheet_as_df() -> pd.DataFrame:
    """
    Lee el rango configurado y devuelve un DataFrame con columnas:
    ['Nombre', 'Asistentes', 'Fecha Confirmación'].
    Si la primera fila son encabezados, se respetan; si no, se asignan.
    """
    if not SHEET_ID:
        raise RuntimeError("Falta SHEET_ID")
    service = _sheets_service()
    resp = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE
    ).execute()
    values = resp.get("values", [])

    if not values:
        return pd.DataFrame(columns=["Nombre", "Asistentes", "Fecha Confirmación"])

    # ¿La primera fila parece encabezado?
    header_candidate = [c.strip().lower() for c in values[0]]
    expected = ["nombre", "asistentes", "fecha confirmación"]
    if all(h in expected for h in header_candidate) and len(values) > 1:
        df = pd.DataFrame(values[1:], columns=[c.strip() for c in values[0]])
    else:
        # Sin encabezado, asignar columnas
        df = pd.DataFrame(values, columns=["Nombre", "Asistentes", "Fecha Confirmación"][:len(values[0])])

    # Normalizaciones
    if "Asistentes" in df.columns:
        df["Asistentes"] = pd.to_numeric(df["Asistentes"], errors="coerce").fillna(0).astype(int)
    if "Fecha Confirmación" in df.columns:
        # no forzar formato; se muestra como texto tal cual fue capturado
        pass

    return df

# --------------------------------------------------------------------
# App FastAPI
# --------------------------------------------------------------------
app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# Cargar datos de Excel (invitados) para validar máximos
df_invitados = pd.read_excel("data/invitados.xlsx")
df_invitados["Grupo"] = df_invitados["Invitación dirigida a"].ffill()
df_invitados["MaxBoletos_num"] = pd.to_numeric(df_invitados["Max Boletos"], errors="coerce").fillna(0).astype(int)

BOLETOS_POR_GRUPO = df_invitados.groupby("Grupo")["MaxBoletos_num"].max().to_dict()
NOMBRES_GRUPO = sorted([g for g in BOLETOS_POR_GRUPO.keys() if isinstance(g, str)])

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("invitacion.html", {
        "request": request,
        "nombres_grupo": NOMBRES_GRUPO,
        "limites_json": json.dumps(BOLETOS_POR_GRUPO, ensure_ascii=False),
        "event_date": "2025-11-15"
    })

@app.post("/confirmar", response_class=HTMLResponse)
async def confirmar(request: Request, nombre: str = Form(...), asistentes: int = Form(...)):
    max_permitido = int(BOLETOS_POR_GRUPO.get(nombre, 0))
    if asistentes > max_permitido:
        return templates.TemplateResponse("invitacion.html", {
            "request": request,
            "nombres_grupo": NOMBRES_GRUPO,
            "limites_json": json.dumps(BOLETOS_POR_GRUPO, ensure_ascii=False),
            "event_date": "2025-11-15",
            "error": f"El máximo permitido para {nombre} es {max_permitido}.",
            "nombre_valor": nombre,
            "asistentes_valor": max_permitido
        }, status_code=400)

    # --- NUEVO: Guardar directamente en Google Sheets ---
    fecha_str = datetime.now(TIMEZONE).strftime("%Y-%m-%d %H:%M:%S")
    try:
        append_to_sheet_row(nombre=nombre, asistentes=int(asistentes), fecha_str=fecha_str)
        saved_to = "google_sheets"
    except Exception as e:
        # Si falla, registramos en logs y como último recurso guardamos CSV local (solo local dev)
        print(f"[CONFIRMACION][ERROR_SHEETS] {e}", flush=True)
        try:
            import csv
            csv_path = "data/confirmaciones.csv"
            existe = os.path.isfile(csv_path)
            with open(csv_path, mode="a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                if not existe:
                    writer.writerow(["Nombre", "Asistentes", "Fecha Confirmación"])
                writer.writerow([nombre, int(asistentes), fecha_str])
            saved_to = "csv_local"
        except Exception as e2:
            print(f"[CONFIRMACION][ERROR_FALLBACK_CSV] {e2}", flush=True)
            saved_to = "none"

    print(f"[CONFIRMACION] nombre={nombre} asistentes={asistentes} saved_to={saved_to} at={datetime.now()}", flush=True)

    return templates.TemplateResponse("gracias.html", {
        "request": request,
        "nombre": nombre,
        "asistentes": asistentes
    })

@app.get("/admin/confirmaciones", response_class=HTMLResponse)
def admin_confirmaciones(request: Request):
    # --- NUEVO: leer directamente desde Google Sheets ---
    try:
        df = read_sheet_as_df()
    except Exception as e:
        print(f"[ADMIN][ERROR_READ_SHEETS] {e}", flush=True)
        # Fallback a CSV local si existiera (sólo local dev)
        csvf = "data/confirmaciones.csv"
        if os.path.exists(csvf):
            df = pd.read_csv(csvf)
        else:
            df = pd.DataFrame(columns=["Nombre", "Asistentes", "Fecha Confirmación"])

    if "Asistentes" in df.columns:
        df["Asistentes"] = pd.to_numeric(df["Asistentes"], errors="coerce").fillna(0).astype(int)

    total_registros = len(df)
    total_asistentes = int(df["Asistentes"].sum()) if "Asistentes" in df.columns else 0

    filas = []
    for _, r in df.fillna("").iterrows():
        filas.append(
            f"<tr><td>{r.get('Nombre','')}</td>"
            f"<td>{r.get('Asistentes','')}</td>"
            f"<td>{r.get('Fecha Confirmación','')}</td></tr>"
        )

    html = f"""
    <html><head><meta charset="utf-8"><title>Confirmaciones</title>
    <style>
      body{{font-family:Arial,sans-serif; padding:20px;}}
      table{{border-collapse:collapse; width:100%;}}
      th,td{{border:1px solid #ddd; padding:8px; text-align:left;}}
      th{{background:#f4f4f4;}}
    </style>
    </head><body>
      <h2>Confirmaciones</h2>
      <p>Total registros: <b>{total_registros}</b> · Total asistentes: <b>{total_asistentes}</b></p>
      <p>
        <a href="/admin/descargar?formato=excel">Descargar Excel</a> ·
        <a href="/admin/descargar?formato=csv">Descargar CSV</a> ·
        <a href="/">Regresar</a>
      </p>
      <table>
        <thead><tr><th>Nombre</th><th>Asistentes</th><th>Fecha</th></tr></thead>
        <tbody>
          {''.join(filas) if filas else '<tr><td colspan="3">Sin registros aún.</td></tr>'}
        </tbody>
      </table>
    </body></html>
    """
    return HTMLResponse(html)

@app.get("/admin/descargar")
def descargar_confirmaciones(formato: str = "excel"):
    # --- NUEVO: descarga a partir de Google Sheets ---
    try:
        df = read_sheet_as_df()
    except Exception as e:
        print(f"[ADMIN][ERROR_READ_SHEETS_FOR_DOWNLOAD] {e}", flush=True)
        df = pd.DataFrame(columns=["Nombre", "Asistentes", "Fecha Confirmación"])

    if formato == "csv":
        csv_bytes = df.to_csv(index=False).encode("utf-8")
        return StreamingResponse(
            io.BytesIO(csv_bytes),
            media_type="text/csv",
            headers={"Content-Disposition": 'attachment; filename="confirmaciones.csv"'}
        )
    else:
        # Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Confirmaciones")
        output.seek(0)
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": 'attachment; filename="confirmaciones.xlsx"'}
        )
